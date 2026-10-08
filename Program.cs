using AccessMacroRunner.Services; // EmailReportService namespace
using Microsoft.Office.Interop.Access;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace AccessMacroRunner
{
    internal class Program
    {
        // ===== Paths / config =====
        private static readonly string DatabasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";
        private static readonly string LogFolder = @"C:\ADSK-Automation\Logs";

        private static readonly string ExportFilePath = Path.Combine(LogFolder, "NetsuiteImportData.csv");
        private static readonly string OldImportFolder = Path.Combine(LogFolder, "Old Imports");

        private static readonly string ExportLogFile = Path.Combine(LogFolder, "ExportAccessIssues.txt");
        private static readonly string MacroLogFile = Path.Combine(LogFolder, "AccessMacroIssues.txt");

        private const string MacroName = "Run Process";
        private const string ExportSourceName = "Netsuite Import Data"; // table or saved query name

        // Global mutex prevents overlapping runs (common cause of double exports / file locks)
        private const string MutexName = @"Global\AccessMacroRunner";

        private static int Main(string[] args)
        {
            string runId = Guid.NewGuid().ToString("N").Substring(0, 8);
            int pid = Process.GetCurrentProcess().Id;

            SafeEnsureDir(LogFolder);

            using var mutex = new Mutex(initiallyOwned: true, name: MutexName, out bool isNew);
            if (!isNew)
            {
                Log(MacroLogFile, runId, pid, "Another instance is already running. Exiting.");
                return 10;
            }

            try
            {
                Log(MacroLogFile, runId, pid, $"Start. Args={(args == null ? "null" : string.Join(" ", args))}");

                if (args != null && args.Length > 0 &&
                    args[0].Equals("emailreports", StringComparison.OrdinalIgnoreCase))
                {
                    Log(MacroLogFile, runId, pid, "Mode=emailreports. Handing off to EmailReportService.");
                    EmailReportService.Run();
                    Log(MacroLogFile, runId, pid, "EmailReportService complete.");
                    return 0;
                }

                if (args != null && args.Length > 0 &&
                    args[0].Equals("twoqtrs", StringComparison.OrdinalIgnoreCase))
                {
                    Log(MacroLogFile, runId, pid, "Mode=twoqtrs. Handing off to TwoQuarterAutomationService.");
                    TwoQuarterAutomationService.Run();
                    Log(MacroLogFile, runId, pid, "TwoQuarterAutomationService complete.");
                    return 0;
                }

                if (args != null && args.Length > 0 &&
                    args[0].Equals("retention", StringComparison.OrdinalIgnoreCase))
                {
                    Log(MacroLogFile, runId, pid, "Mode=retention. Handing off to RetentionService.");
                    RetentionService.Run();
                    Log(MacroLogFile, runId, pid, "RetentionService complete.");
                    return 0;
                }

                RunMacroAndExport(runId, pid);
                Log(MacroLogFile, runId, pid, "Complete.");
                return 0;
            }
            catch (Exception ex)
            {
                Log(MacroLogFile, runId, pid, $"FATAL: {ex}");
                return 1;
            }
        }

        private static void RunMacroAndExport(string runId, int pid)
        {
            if (!File.Exists(DatabasePath))
                throw new FileNotFoundException("Access database not found.", DatabasePath);

            SafeEnsureDir(LogFolder);
            SafeEnsureDir(OldImportFolder);

            var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");

            // 1) Archive old export first (fail fast if we can't)
            string datedArchiveFolder = Path.Combine(
                OldImportFolder,
                DateTime.Now.ToString("yyyy"),
                DateTime.Now.ToString("MM"));
            ArchiveIfExists(ExportFilePath, datedArchiveFolder, timestamp, ExportLogFile, runId, pid);

            // 2) Open Access, run macro, export using Access engine (TransferText)
            Application accessApp = null;
            try
            {
                Log(MacroLogFile, runId, pid, $"Opening Access DB: {DatabasePath}");
                accessApp = new Application();
                accessApp.OpenCurrentDatabase(DatabasePath, false);

                Log(MacroLogFile, runId, pid, $"Running macro: {MacroName}");
                accessApp.DoCmd.RunMacro(MacroName);
                Log(MacroLogFile, runId, pid, $"Macro executed: {MacroName}");

                Log(ExportLogFile, runId, pid, $"Exporting '{ExportSourceName}' -> {ExportFilePath}");
                ExportTableOrQueryToCsv_AccessCom(accessApp, ExportSourceName, ExportFilePath, ExportLogFile, runId, pid);

                // 3) Clean midnight timestamps
                CleanCsvOfMidnight(ExportFilePath, ExportLogFile, runId, pid);
            }
            finally
            {
                SafeCloseAccess(ref accessApp, MacroLogFile, runId, pid);
            }
        }

        /// <summary>
        /// Export via Access COM (DoCmd.TransferText). More reliable than OleDb Text driver SELECT INTO.
        /// </summary>
        private static void ExportTableOrQueryToCsv_AccessCom(
            Application accessApp,
            string tableOrQueryName,
            string exportFilePath,
            string logFile,
            string runId,
            int pid)
        {
            try
            {
                // If something already exists (shouldn't, because we archive first), remove it to avoid prompt/overwrite issues.
                if (File.Exists(exportFilePath))
                {
                    File.Delete(exportFilePath);
                    Log(logFile, runId, pid, $"Deleted existing export to allow new write: {exportFilePath}");
                }

                accessApp.DoCmd.TransferText(
                    AcTextTransferType.acExportDelim,
                    Type.Missing,         // export spec (none)
                    tableOrQueryName,     // table/query name
                    exportFilePath,       // output path
                    true                  // include headers
                );

                if (!File.Exists(exportFilePath))
                    throw new IOException($"Expected export file not found after export: {exportFilePath}");

                long len = new FileInfo(exportFilePath).Length;
                Log(logFile, runId, pid, $"Exported '{tableOrQueryName}' ({len:n0} bytes).");
            }
            catch (Exception ex)
            {
                Log(logFile, runId, pid, $"ERROR exporting '{tableOrQueryName}': {ex.Message}");
                if (ex.InnerException != null)
                    Log(logFile, runId, pid, $"Inner: {ex.InnerException.Message}");
                throw;
            }
        }

        private static void ArchiveIfExists(
            string exportFilePath,
            string archiveFolder,
            string timestamp,
            string logFile,
            string runId,
            int pid)
        {
            if (!File.Exists(exportFilePath))
            {
                Log(logFile, runId, pid, "No previous export found to archive.");
                return;
            }

            SafeEnsureDir(archiveFolder);

            string archivedPath = Path.Combine(
                archiveFolder,
                $"{Path.GetFileNameWithoutExtension(exportFilePath)}_{timestamp}{Path.GetExtension(exportFilePath)}"
            );

            try
            {
                File.Move(exportFilePath, archivedPath);
                Log(logFile, runId, pid, $"Archived previous export to: {archivedPath}");
            }
            catch (IOException)
            {
                // Fallback if file system refuses atomic move (e.g., cross-volume or transient locks)
                try
                {
                    File.Copy(exportFilePath, archivedPath, overwrite: true);
                    File.Delete(exportFilePath);
                    Log(logFile, runId, pid, $"Archived previous export to: {archivedPath} (copy/delete fallback)");
                }
                catch (Exception ex)
                {
                    Log(logFile, runId, pid, $"ERROR archiving previous export: {ex.Message}");
                    if (ex.InnerException != null)
                        Log(logFile, runId, pid, $"Inner: {ex.InnerException.Message}");
                    throw;
                }
            }
        }

        private static void CleanCsvOfMidnight(string path, string logFile, string runId, int pid)
        {
            try
            {
                if (!File.Exists(path))
                {
                    Log(logFile, runId, pid, $"CSV not found for cleaning: {path}");
                    return;
                }

                // Only remove midnight time components; keep other times intact
                string text = File.ReadAllText(path, Encoding.UTF8);

                // Examples:
                // "11/17/2025 0:00:00"  -> "11/17/2025"
                // "11/17/2025 00:00:00" -> "11/17/2025"
                text = Regex.Replace(text, @"\s0{1,2}:00:00\b", "");

                File.WriteAllText(path, text, Encoding.UTF8);
                Log(logFile, runId, pid, "Cleaned CSV of midnight timestamps.");
            }
            catch (Exception ex)
            {
                Log(logFile, runId, pid, $"ERROR scrubbing CSV: {ex.Message}");
                if (ex.InnerException != null)
                    Log(logFile, runId, pid, $"Inner: {ex.InnerException.Message}");
                // do not throw; cleaning is non-fatal
            }
        }

        private static void SafeCloseAccess(ref Application accessApp, string logFile, string runId, int pid)
        {
            try
            {
                if (accessApp == null) return;

                try { accessApp.CloseCurrentDatabase(); } catch { }
                try { accessApp.Quit(); } catch { }

                try
                {
                    Marshal.FinalReleaseComObject(accessApp);
                }
                catch
                {
                    // If FinalRelease fails for some reason, fall back (best-effort)
                    try { Marshal.ReleaseComObject(accessApp); } catch { }
                }

                accessApp = null;
            }
            catch (Exception ex)
            {
                Log(logFile, runId, pid, $"Cleanup error: {ex.Message}");
                if (ex.InnerException != null)
                    Log(logFile, runId, pid, $"Inner: {ex.InnerException.Message}");
            }
            finally
            {
                // Ensure COM is fully collected
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void SafeEnsureDir(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }
            catch
            {
                // swallow; caller logs elsewhere
            }
        }

        private static void Log(string file, string runId, int pid, string message)
        {
            try
            {
                SafeEnsureDir(Path.GetDirectoryName(file) ?? ".");
                File.AppendAllText(
                    file,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - RunId={runId} PID={pid} - {message}{Environment.NewLine}"
                );
            }
            catch
            {
                // swallow logging errors
            }
        }
    }
}

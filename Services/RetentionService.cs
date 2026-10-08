using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace AccessMacroRunner.Services
{
    internal static class RetentionService
    {
        private static readonly string[] RunnerLogNames =
        {
            "AccessMacroIssues.txt",
            "ExportAccessIssues.txt",
            "EmailReportLog.txt",
            "AccessMacroVerification.txt",
            "TwoQuarterAutomationLog.txt"
        };

        private sealed class Summary
        {
            public int Examined;
            public int WouldStage;
            public int Staged;
            public int Purged;
            public int Errors;
            public long Bytes;
        }

        public static void Run()
        {
            var settings = Settings.Load();
            string logs = Path.Combine(settings.Root, "Logs");
            string retentionLog = Path.Combine(logs, "Retention", "RetentionLog_" + DateTime.Now.ToString("yyyy-MM") + ".txt");
            var summary = new Summary();

            EnsureDirectory(Path.GetDirectoryName(retentionLog));
            Log(retentionLog, "Starting retention. DryRun=" + settings.DryRun + ", Root=" + settings.Root + ".");

            if (!settings.Enabled)
            {
                Log(retentionLog, "RetentionEnabled is false; no files were examined or changed.");
                return;
            }

            RotateRunnerLogs(settings, retentionLog, summary);
            SnapshotRawDownloads(settings, retentionLog, summary);

            ApplyRule(Path.Combine(logs, "Archive"), "logs", DateTime.Now.AddDays(-settings.LogDays),
                new[] { ".txt", ".log" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(logs, "Retention"), "logs", DateTime.Now.AddDays(-settings.LogDays),
                new[] { ".txt", ".log" }, settings, retentionLog, summary, retentionLog);
            ApplyRule(Path.Combine(logs, "Old Imports"), "generated-exports", DateTime.Now.AddDays(-settings.GeneratedExportDays),
                new[] { ".csv" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "NetSuiteImports"), "netsuite-inputs", DateTime.Now.AddDays(-settings.NetSuiteInputDays),
                new[] { ".csv" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "FormaDownloads", "Archive"), "raw-downloads", DateTime.Now.AddDays(-settings.RawDownloadDays),
                new[] { ".csv", ".gz", ".json" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "2Qtrs", "Archive"), "two-quarter-inputs", DateTime.Now.AddDays(-settings.TwoQuarterDays),
                new[] { ".csv" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "2Qtrs", "Old"), "two-quarter-inputs", DateTime.Now.AddDays(-settings.TwoQuarterDays),
                new[] { ".csv" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "EmailReports"), "email-reports", DateTime.Now.AddMonths(-settings.EmailReportMonths),
                new[] { ".xlsx" }, settings, retentionLog, summary);
            ApplyRule(Path.Combine(settings.Root, "FailedRuns"), "failed-artifacts", DateTime.Now.AddDays(-settings.FailedArtifactDays),
                new[] { ".csv", ".gz", ".json", ".xlsx", ".txt", ".log" }, settings, retentionLog, summary);

            OrganizeLegacyFiles(Path.Combine(settings.Root, "EmailReports"), new[] { ".xlsx" }, settings, retentionLog, summary);
            OrganizeLegacyFiles(Path.Combine(logs, "Old Imports"), new[] { ".csv" }, settings, retentionLog, summary);
            PurgeQuarantine(settings, retentionLog, summary);

            Log(retentionLog,
                "Complete. Examined=" + summary.Examined
                + ", WouldStage=" + summary.WouldStage
                + ", Staged=" + summary.Staged
                + ", Purged=" + summary.Purged
                + ", Bytes=" + summary.Bytes
                + ", Errors=" + summary.Errors + ".");
        }

        private static void ApplyRule(
            string sourceRoot,
            string category,
            DateTime cutoff,
            string[] extensions,
            Settings settings,
            string logPath,
            Summary summary,
            string excludedFile = null)
        {
            if (!Directory.Exists(sourceRoot))
                return;

            foreach (string file in SafeFiles(sourceRoot, logPath, summary))
            {
                if (excludedFile != null && PathsEqual(file, excludedFile))
                    continue;
                if (!extensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                    continue;

                summary.Examined++;
                try
                {
                    var info = new FileInfo(file);
                    DateTime effectiveDate = GetEffectiveDate(info);
                    if (effectiveDate >= cutoff)
                        continue;

                    if (settings.DryRun)
                    {
                        summary.WouldStage++;
                        summary.Bytes += info.Length;
                        Log(logPath, "DRY RUN: Would stage " + file + " (effective date " + effectiveDate.ToString("yyyy-MM-dd") + ").");
                        continue;
                    }

                    string quarantine = Path.Combine(settings.Root, "PendingDeletion", category, DateTime.Now.ToString("yyyy-MM-dd"));
                    EnsureDirectory(quarantine);
                    string destination = UniquePath(Path.Combine(quarantine, Path.GetFileName(file)));
                    File.Move(file, destination);
                    File.SetLastWriteTime(destination, DateTime.Now);
                    summary.Staged++;
                    summary.Bytes += info.Length;
                    Log(logPath, "Staged for deletion: " + file + " -> " + destination + ".");
                }
                catch (Exception ex)
                {
                    summary.Errors++;
                    Log(logPath, "ERROR processing " + file + ": " + ex.Message);
                }
            }
        }

        private static void RotateRunnerLogs(Settings settings, string logPath, Summary summary)
        {
            string logs = Path.Combine(settings.Root, "Logs");
            string archive = Path.Combine(logs, "Archive");
            string marker = Path.Combine(archive, ".last-rotation-month");
            string currentMonth = DateTime.Now.ToString("yyyy-MM");

            try
            {
                string priorMonth = File.Exists(marker) ? File.ReadAllText(marker).Trim() : string.Empty;
                if (string.IsNullOrWhiteSpace(priorMonth))
                {
                    if (settings.DryRun)
                        Log(logPath, "DRY RUN: Would initialize monthly log rotation marker to " + currentMonth + ".");
                    else
                    {
                        EnsureDirectory(archive);
                        File.WriteAllText(marker, currentMonth);
                        Log(logPath, "Initialized monthly log rotation marker to " + currentMonth + ".");
                    }
                    return;
                }

                if (priorMonth.Equals(currentMonth, StringComparison.OrdinalIgnoreCase))
                    return;

                DateTime archiveMonth;
                if (!DateTime.TryParseExact(priorMonth, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out archiveMonth))
                    archiveMonth = DateTime.Now.AddMonths(-1);

                string destinationFolder = Path.Combine(archive, archiveMonth.ToString("yyyy"), archiveMonth.ToString("MM"));
                foreach (string name in RunnerLogNames)
                {
                    string source = Path.Combine(logs, name);
                    if (!File.Exists(source))
                        continue;

                    string destination = UniquePath(Path.Combine(
                        destinationFolder,
                        Path.GetFileNameWithoutExtension(name) + "_through_" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + Path.GetExtension(name)));
                    if (settings.DryRun)
                        Log(logPath, "DRY RUN: Would rotate " + source + " -> " + destination + ".");
                    else
                    {
                        EnsureDirectory(destinationFolder);
                        File.Move(source, destination);
                        Log(logPath, "Rotated " + source + " -> " + destination + ".");
                    }
                }

                if (!settings.DryRun)
                    File.WriteAllText(marker, currentMonth);
            }
            catch (Exception ex)
            {
                summary.Errors++;
                Log(logPath, "ERROR rotating runner logs: " + ex.Message);
            }
        }

        private static void SnapshotRawDownloads(Settings settings, string logPath, Summary summary)
        {
            string downloadFolder = Path.Combine(settings.Root, "FormaDownloads");
            string[] activeNames = { "all_transaction_fees.csv", "all_transaction_fees.csv.gz" };

            foreach (string name in activeNames)
            {
                string source = Path.Combine(downloadFolder, name);
                if (!File.Exists(source))
                    continue;

                try
                {
                    var info = new FileInfo(source);
                    string archiveFolder = Path.Combine(
                        downloadFolder,
                        "Archive",
                        info.LastWriteTime.ToString("yyyy"),
                        info.LastWriteTime.ToString("MM"));
                    string destination = Path.Combine(
                        archiveFolder,
                        Path.GetFileNameWithoutExtension(name) + "_" + info.LastWriteTime.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(name));

                    // A stable name based on the source modification time prevents duplicate daily copies.
                    if (File.Exists(destination))
                        continue;

                    if (settings.DryRun)
                    {
                        Log(logPath, "DRY RUN: Would snapshot active raw download " + source + " -> " + destination + ".");
                    }
                    else
                    {
                        EnsureDirectory(archiveFolder);
                        File.Copy(source, destination, false);
                        File.SetLastWriteTime(destination, info.LastWriteTime);
                        Log(logPath, "Snapshotted active raw download " + source + " -> " + destination + ".");
                    }
                }
                catch (Exception ex)
                {
                    summary.Errors++;
                    Log(logPath, "ERROR snapshotting raw download " + source + ": " + ex.Message);
                }
            }
        }

        private static void OrganizeLegacyFiles(
            string folder,
            string[] extensions,
            Settings settings,
            string logPath,
            Summary summary)
        {
            if (!Directory.Exists(folder))
                return;

            foreach (string file in Directory.GetFiles(folder, "*", SearchOption.TopDirectoryOnly))
            {
                if (!extensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase))
                    continue;

                try
                {
                    DateTime date = GetEffectiveDate(new FileInfo(file));
                    string destinationFolder = Path.Combine(folder, date.ToString("yyyy"), date.ToString("MM"));
                    string destination = UniquePath(Path.Combine(destinationFolder, Path.GetFileName(file)));
                    if (settings.DryRun)
                        Log(logPath, "DRY RUN: Would organize " + file + " -> " + destination + ".");
                    else
                    {
                        EnsureDirectory(destinationFolder);
                        File.Move(file, destination);
                        Log(logPath, "Organized " + file + " -> " + destination + ".");
                    }
                }
                catch (Exception ex)
                {
                    summary.Errors++;
                    Log(logPath, "ERROR organizing " + file + ": " + ex.Message);
                }
            }
        }

        private static void PurgeQuarantine(Settings settings, string logPath, Summary summary)
        {
            string quarantine = Path.Combine(settings.Root, "PendingDeletion");
            if (!Directory.Exists(quarantine))
                return;

            DateTime cutoff = DateTime.Now.AddDays(-settings.QuarantineDays);
            foreach (string file in SafeFiles(quarantine, logPath, summary))
            {
                try
                {
                    var info = new FileInfo(file);
                    if (info.LastWriteTime >= cutoff)
                        continue;

                    if (settings.DryRun)
                    {
                        Log(logPath, "DRY RUN: Would permanently delete quarantined file " + file + ".");
                        continue;
                    }

                    long length = info.Length;
                    File.Delete(file);
                    summary.Purged++;
                    summary.Bytes += length;
                    Log(logPath, "Permanently deleted quarantined file " + file + ".");
                }
                catch (Exception ex)
                {
                    summary.Errors++;
                    Log(logPath, "ERROR purging " + file + ": " + ex.Message);
                }
            }
        }

        private static IEnumerable<string> SafeFiles(string folder, string logPath, Summary summary)
        {
            try
            {
                return Directory.GetFiles(folder, "*", SearchOption.AllDirectories);
            }
            catch (Exception ex)
            {
                summary.Errors++;
                Log(logPath, "ERROR enumerating " + folder + ": " + ex.Message);
                return Enumerable.Empty<string>();
            }
        }

        private static string UniquePath(string path)
        {
            if (!File.Exists(path))
                return path;

            return Path.Combine(
                Path.GetDirectoryName(path),
                Path.GetFileNameWithoutExtension(path) + "_" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension(path));
        }

        private static DateTime GetEffectiveDate(FileInfo file)
        {
            string name = Path.GetFileNameWithoutExtension(file.Name);
            string[] patterns =
            {
                @"(?<!\d)(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})(?!\d)",
                @"(?<!\d)(\d{4}-\d{2}-\d{2}_\d{6})(?!\d)",
                @"(?<!\d)(\d{8}_\d{6})(?!\d)",
                @"(?<!\d)(\d{1,2}-\d{1,2}-\d{2})(?!\d)"
            };
            string[] formats =
            {
                "yyyy-MM-dd_HH-mm-ss",
                "yyyy-MM-dd_HHmmss",
                "yyyyMMdd_HHmmss",
                "M-d-yy"
            };

            for (int i = 0; i < patterns.Length; i++)
            {
                Match match = Regex.Match(name, patterns[i]);
                DateTime parsed;
                if (match.Success && DateTime.TryParseExact(
                    match.Groups[1].Value,
                    formats[i],
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out parsed))
                {
                    return parsed;
                }
            }

            return file.LastWriteTime;
        }

        private static bool PathsEqual(string left, string right)
        {
            return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }

        private static void EnsureDirectory(string path)
        {
            if (!string.IsNullOrWhiteSpace(path) && !Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private static void Log(string path, string message)
        {
            try
            {
                EnsureDirectory(Path.GetDirectoryName(path));
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine);
            }
            catch
            {
                // Cleanup remains non-fatal even when its own log cannot be written.
            }
        }

        private sealed class Settings
        {
            public string Root;
            public bool Enabled;
            public bool DryRun;
            public int LogDays;
            public int NetSuiteInputDays;
            public int GeneratedExportDays;
            public int RawDownloadDays;
            public int TwoQuarterDays;
            public int EmailReportMonths;
            public int FailedArtifactDays;
            public int QuarantineDays;

            public static Settings Load()
            {
                string root = Path.GetFullPath(GetString("AutomationRoot", @"C:\ADSK-Automation"));
                if (string.Equals(
                    root.TrimEnd(Path.DirectorySeparatorChar),
                    Path.GetPathRoot(root).TrimEnd(Path.DirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase))
                {
                    throw new ConfigurationErrorsException("AutomationRoot cannot be a drive root.");
                }

                return new Settings
                {
                    Root = root,
                    Enabled = GetBool("RetentionEnabled", true),
                    DryRun = GetBool("RetentionDryRun", true),
                    LogDays = GetPositiveInt("RetentionLogDays", 180),
                    NetSuiteInputDays = GetPositiveInt("RetentionNetSuiteInputDays", 90),
                    GeneratedExportDays = GetPositiveInt("RetentionGeneratedExportDays", 180),
                    RawDownloadDays = GetPositiveInt("RetentionRawDownloadDays", 30),
                    TwoQuarterDays = GetPositiveInt("RetentionTwoQuarterDays", 180),
                    EmailReportMonths = GetPositiveInt("RetentionEmailReportMonths", 13),
                    FailedArtifactDays = GetPositiveInt("RetentionFailedArtifactDays", 180),
                    QuarantineDays = GetPositiveInt("RetentionQuarantineDays", 7)
                };
            }

            private static string GetString(string key, string fallback)
            {
                string value = ConfigurationManager.AppSettings[key];
                return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
            }

            private static bool GetBool(string key, bool fallback)
            {
                bool value;
                return bool.TryParse(ConfigurationManager.AppSettings[key], out value) ? value : fallback;
            }

            private static int GetPositiveInt(string key, int fallback)
            {
                int value;
                return int.TryParse(ConfigurationManager.AppSettings[key], out value) && value > 0 ? value : fallback;
            }
        }
    }
}

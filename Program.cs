using AccessMacroRunner.Services;  // <-- your EmailReportService namespace 
using Microsoft.Office.Interop.Access;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessMacroRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 &&
                args[0].Equals("emailreports", StringComparison.OrdinalIgnoreCase))
            {
                EmailReportService.Run();
                return;
            }

            RunMacroAndExport();
        }

        private static void RunMacroAndExport()
        {
            string databasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";
            string logFolder = @"C:\ADSK-Automation\Logs";
            string exportFilePath = Path.Combine(logFolder, "NetsuiteImportData.csv");
            string oldImportFolder = Path.Combine(logFolder, "Old Imports");
            string exportLogFile = Path.Combine(logFolder, "ExportAccessIssues.txt");
            string macroLogFile = Path.Combine(logFolder, "AccessMacroIssues.txt");
            string macroName = "Run Process";
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");

            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
            Application accessApp = null;

            try
            {
                Log(exportLogFile, "🚀 Starting RunMacroAndExport");

                Directory.CreateDirectory(logFolder);
                Log(macroLogFile, $"🔄 Launching Access DB: {databasePath}");

                accessApp = new Application();
                accessApp.OpenCurrentDatabase(databasePath);
                accessApp.DoCmd.RunMacro(macroName);
                Log(macroLogFile, $"✅ Macro '{macroName}' executed.");

                // Move old file if it exists
                Log(exportLogFile, $"🔍 Checking for existing export file: {exportFilePath}");
                if (File.Exists(exportFilePath))
                {
                    Log(exportLogFile, $"📁 Old export file exists, preparing to move...");
                    Directory.CreateDirectory(oldImportFolder);

                    string archivedPath = Path.Combine(oldImportFolder, $"NetsuiteImportData_{timestamp}.csv");
                    Log(exportLogFile, $"📦 Moving file to archive: {archivedPath}");

                    try
                    {
                        File.Move(exportFilePath, archivedPath);
                        Log(exportLogFile, $"✅ Moved old export to: {archivedPath}");
                    }
                    catch (Exception moveEx)
                    {
                        Log(exportLogFile, $"❌ Failed to move old export: {moveEx.Message}");
                        if (moveEx.InnerException != null)
                            Log(exportLogFile, $"   ⤷ Inner: {moveEx.InnerException.Message}");
                    }
                }
                else
                {
                    Log(exportLogFile, $"ℹ️ No previous export file found.");
                }

                // Export Access table to CSV
                Log(exportLogFile, $"🔄 Connecting to Access DB to export new CSV...");

                using (var connection = new OleDbConnection(connStr))
                {
                    connection.Open();
                    Log(exportLogFile, $"✅ Access DB connection opened.");

                    // Log column names
                    try
                    {
                        using (var schemaCmd = new OleDbCommand("SELECT * FROM [Netsuite Import Data]", connection))
                        using (var reader = schemaCmd.ExecuteReader(CommandBehavior.SchemaOnly))
                        {
                            var schemaTable = reader.GetSchemaTable();
                            if (schemaTable != null)
                            {
                                var columns = schemaTable.Rows.Cast<DataRow>()
                                    .Select(row => row["ColumnName"].ToString())
                                    .ToArray();

                                Log(exportLogFile, $"📋 Export columns: {string.Join(", ", columns)}");
                            }
                            else
                            {
                                Log(exportLogFile, "⚠️ Schema table was null.");
                            }
                        }
                    }
                    catch (Exception schemaEx)
                    {
                        Log(exportLogFile, $"❌ Failed to read schema: {schemaEx.Message}");
                    }

                    // Execute export query
                    string exportQuery = $@"
            SELECT * 
            INTO [Text;FMT=Delimited;HDR=Yes;Database={logFolder};].[NetsuiteImportData.csv] 
            FROM [Netsuite Import Data]";

                    Log(exportLogFile, $"📤 Running export query:\n{exportQuery}");

                    try
                    {
                        using (var command = new OleDbCommand(exportQuery, connection))
                        {
                            command.ExecuteNonQuery();
                            Log(exportLogFile, $"✅ Exported Access table to: {exportFilePath}");
                        }
                    }
                    catch (Exception exportEx)
                    {
                        Log(exportLogFile, $"❌ Export failed: {exportEx.Message}");
                    }
                }

                CleanCsvOfMidnight(exportFilePath, exportLogFile);
            }
            catch (Exception ex)
            {
                Log(macroLogFile, $"❌ ERROR: {ex.Message}");
                if (ex.InnerException != null)
                    Log(macroLogFile, $"   ⤷ Inner: {ex.InnerException.Message}");
            }
            finally
            {
                try
                {
                    accessApp?.CloseCurrentDatabase();
                    accessApp?.Quit();
                    if (accessApp != null)
                    {
                        Marshal.ReleaseComObject(accessApp);
                        accessApp = null;
                    }
                }
                catch (Exception ex)
                {
                    Log(macroLogFile, $"⚠️ Cleanup Error: {ex.Message}");
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        static void CleanCsvOfMidnight(string path, string log)
        {
            try
            {
                var lines = File.ReadAllLines(path);
                for (int i = 0; i < lines.Length; i++)
                    lines[i] = lines[i].Replace(" 0:00:00", "");
                File.WriteAllLines(path, lines);
                Log(log, $"🧼 Cleaned CSV of ' 0:00:00'.");
            }
            catch (Exception ex)
            {
                Log(log, $"❌ Error scrubbing CSV: {ex.Message}");
            }
        }

        static void Log(string file, string message)
        {
            File.AppendAllText(file, $"{DateTime.Now:yyyy-MM-dd HH:mm} - {message}{Environment.NewLine}");
        }
    }
}

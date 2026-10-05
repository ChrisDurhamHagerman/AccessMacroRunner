using ClosedXML.Excel;
using Microsoft.Office.Interop.Access;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace AccessMacroRunner.Services
{
    public static class EmailReportService
    {
        private static readonly string _dbPath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";
        private static readonly string _logPath = @"C:\ADSK-Automation\Logs\EmailReportLog.txt";
        private static readonly string _outputDir = @"C:\ADSK-Automation\EmailReports";
        private static readonly EmailSettings _emailSettings = EmailSettings.Load();

        public static void Run()
        {
            try
            {
                EnsureDirectory(Path.GetDirectoryName(_logPath));
                EnsureDirectory(_outputDir);

                // Clean output dir
                foreach (var file in Directory.GetFiles(_outputDir))
                {
                    try { File.Delete(file); }
                    catch { /* ignore */ }
                }

                SendReport(
                    queryName: "Qry-NBE RA Orders",
                    subject: "Autodesk NBE RAs",
                    body: "Here is a list of RAs from Autodesk.",
                    fileNameBase: "NBE RA Orders",
                    to: new[] { "SSG-Reports@hagerman.com" },
                    cc: new[] { "DavidHagerman@hagerman.com", "ChrisDurham@hagerman.com" }
                );

                SendReport(
                    queryName: "Qry-NBE Orders without Quotes",
                    subject: "Autodesk NBE Orders without Quotes.",
                    body: "Here is a list of Autodesk NBE Orders without Quotes.",
                    fileNameBase: "NBE Orders without Quote",
                    to: new[] { "JeremyStefanek@hagerman.com", "SSG-Reports@hagerman.com" },
                    cc: new[] { "DavidHagerman@hagerman.com", "ChrisDurham@hagerman.com" }
                );

                RunVerificationMacro();
            }
            catch (Exception ex)
            {
                AppendLog(_logPath, "Fatal Error: " + ex);
            }
        }

        private static void SendReport(
            string queryName,
            string subject,
            string body,
            string fileNameBase,
            string[] to,
            string[] cc)
        {
            try
            {
                string query = "SELECT * FROM [" + queryName + "]";
                string timestamp = DateTime.Now.ToString("M-d-yy");
                string filePath = Path.Combine(_outputDir, fileNameBase + " " + timestamp + ".xlsx");

                ExportQueryToExcel(query, filePath);
                SendEmail(subject, body, filePath, to, cc);

                AppendLog(_logPath, subject + " sent to " + string.Join(",", to));
            }
            catch (Exception ex)
            {
                AppendLog(_logPath, "Error generating '" + subject + "': " + ex);
            }
        }

        private static void ExportQueryToExcel(string query, string filePath)
        {
            string connStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _dbPath + ";Persist Security Info=False;";

            using (var connection = new OleDbConnection(connStr))
            using (var command = new OleDbCommand(query, connection))
            using (var adapter = new OleDbDataAdapter(command))
            {
                var table = new DataTable();
                connection.Open();
                adapter.Fill(table);

                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Results");

                    // Headers
                    for (int c = 0; c < table.Columns.Count; c++)
                        ws.Cell(1, c + 1).SetValue(table.Columns[c].ColumnName);

                    var headerRow = ws.Row(1);
                    headerRow.Style.Font.Bold = false;
                    headerRow.Style.Fill.PatternType = XLFillPatternValues.Solid;
                    headerRow.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Background1, -0.25);
                    ws.Columns(1, table.Columns.Count).AdjustToContents(1, 1);

                    // Data rows
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        var row = table.Rows[r];
                        for (int c = 0; c < table.Columns.Count; c++)
                        {
                            var obj = row[c];
                            var text = obj == DBNull.Value ? "" : obj.ToString();
                            ws.Cell(r + 2, c + 1).SetValue(text);
                        }
                    }

                    var used = ws.RangeUsed();
                    if (used != null)
                    {
                        used.Style.Alignment.WrapText = true;
                        used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    }

                    workbook.SaveAs(filePath);
                }
            }
        }

        private static void SendEmail(
            string subject,
            string bodyText,
            string attachmentPath,
            string[] to,
            string[] cc)
        {
            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException("Attachment not found", attachmentPath);

            var finalTo = to ?? Array.Empty<string>();
            var finalCc = cc ?? Array.Empty<string>();
            var finalSubject = subject;
            var finalBody = bodyText;

            if (_emailSettings.DebugMode)
            {
                var originalTo = string.Join(", ", finalTo);
                var originalCc = string.Join(", ", finalCc);

                finalTo = new[] { _emailSettings.DebugRecipient };
                finalCc = Array.Empty<string>();
                finalSubject = "[DEBUG] " + subject;
                finalBody = bodyText
                    + Environment.NewLine
                    + Environment.NewLine
                    + "Debug mode is enabled. Original To: " + originalTo
                    + Environment.NewLine
                    + "Original CC: " + originalCc;

                AppendLog(_logPath, "Email debug mode enabled. Redirecting '" + subject + "' to " + _emailSettings.DebugRecipient + ".");
            }

            using (var smtp = new SmtpClient(_emailSettings.SmtpHost, _emailSettings.SmtpPort)
            {
                EnableSsl = _emailSettings.EnableSsl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false
            })
            using (var mail = new MailMessage())
            {
                mail.From = new MailAddress(_emailSettings.FromAddress, _emailSettings.FromDisplayName);
                mail.Subject = finalSubject;
                mail.Body = finalBody;
                mail.IsBodyHtml = false;

                foreach (var addr in finalTo.Where(IsValidAddress)) mail.To.Add(addr);
                foreach (var addr in finalCc.Where(IsValidAddress)) mail.CC.Add(addr);

                mail.Attachments.Add(new System.Net.Mail.Attachment(attachmentPath));

                smtp.Send(mail);
            }
        }

        private static void RunVerificationMacro()
        {
            string logDir = Path.GetDirectoryName(_logPath);
            if (string.IsNullOrWhiteSpace(logDir))
                logDir = @"C:\ADSK-Automation\Logs";

            EnsureDirectory(logDir);

            var macroLogFile = Path.Combine(logDir, "AccessMacroVerification.txt");

            Application accessApp = null;
            try
            {
                AppendLog(macroLogFile, "Running verification macro on database: " + _dbPath);

                accessApp = new Application();
                accessApp.OpenCurrentDatabase(_dbPath, false);
                accessApp.DoCmd.RunMacro("Run Process AFTER VERIFICATION");

                AppendLog(macroLogFile, "Verification macro completed successfully.");
            }
            catch (Exception ex)
            {
                AppendLog(macroLogFile, "Error running verification macro: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (accessApp != null)
                    {
                        try { accessApp.CloseCurrentDatabase(); } catch { }
                        try { accessApp.Quit(); } catch { }
                        try { Marshal.FinalReleaseComObject(accessApp); } catch { }
                        accessApp = null;
                    }
                }
                catch { /* ignore */ }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void EnsureDirectory(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
        }

        private static void AppendLog(string filePath, string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(filePath);
                EnsureDirectory(dir);
                File.AppendAllText(filePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " - " + message + Environment.NewLine);
            }
            catch
            {
                // swallow logging errors
            }
        }

        private static bool IsValidAddress(string address)
        {
            return !string.IsNullOrWhiteSpace(address);
        }

        private sealed class EmailSettings
        {
            public bool DebugMode { get; private set; }
            public string DebugRecipient { get; private set; }
            public string SmtpHost { get; private set; }
            public int SmtpPort { get; private set; }
            public bool EnableSsl { get; private set; }
            public string FromAddress { get; private set; }
            public string FromDisplayName { get; private set; }

            public static EmailSettings Load()
            {
                return new EmailSettings
                {
                    DebugMode = GetBool("EmailDebugMode", false),
                    DebugRecipient = GetString("EmailDebugRecipient", "chrisdurham@hagerman.com"),
                    SmtpHost = GetString("SmtpHost", "d250495a.ess.barracudanetworks.com"),
                    SmtpPort = GetInt("SmtpPort", 587),
                    EnableSsl = GetBool("SmtpEnableSsl", false),
                    FromAddress = GetString("MailFromAddress", "no-reply@hagerman.com"),
                    FromDisplayName = GetString("MailFromDisplayName", "Hagerman & Company")
                };
            }

            private static string GetString(string key, string defaultValue)
            {
                var value = ConfigurationManager.AppSettings[key];
                return string.IsNullOrWhiteSpace(value) ? defaultValue : value.Trim();
            }

            private static bool GetBool(string key, bool defaultValue)
            {
                var value = ConfigurationManager.AppSettings[key];
                return bool.TryParse(value, out var result) ? result : defaultValue;
            }

            private static int GetInt(string key, int defaultValue)
            {
                var value = ConfigurationManager.AppSettings[key];
                return int.TryParse(value, out var result) ? result : defaultValue;
            }
        }
    }
}

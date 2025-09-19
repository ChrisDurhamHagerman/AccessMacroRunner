using ClosedXML.Excel;
using Microsoft.Office.Interop.Access;
using System;
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

        public static void Run()
        {
            try
            {
                if (Directory.Exists(_outputDir))
                {
                    foreach (var file in Directory.GetFiles(_outputDir))
                        File.Delete(file);
                }
                else
                {
                    Directory.CreateDirectory(_outputDir);
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

                //SendReport(
                //    queryName: "Qry-NBE RA Orders",
                //    subject: "TEST Autodesk NBE RAs",
                //    body: "TEST Here is a list of RAs from Autodesk.",
                //    fileNameBase: "NBE RA Orders",
                //    to: new[] { "ChrisDurham@hagerman.com" },
                //    cc: Array.Empty<string>()
                //);

                //SendReport(
                //    queryName: "Qry-NBE Orders without Quotes",
                //    subject: "TEST Autodesk NBE Orders without Quotes.",
                //    body: "TEST Here is a list of Autodesk NBE Orders without Quotes.",
                //    fileNameBase: "NBE Orders without Quote",
                //    to: new[] { "ChrisDurham@hagerman.com" },
                //    cc: Array.Empty<string>()
                //);

                RunVerificationMacro();
            }
            catch (Exception ex)
            {
                File.AppendAllText(
                    _logPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm} - ❌ Fatal Error: {ex.Message}{Environment.NewLine}"
                );
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
                string query = $"SELECT * FROM [{queryName}]";
                string timestamp = DateTime.Now.ToString("M-d-yy");
                string filePath = Path.Combine(_outputDir, $"{fileNameBase} {timestamp}.xlsx");

                ExportQueryToExcel(query, filePath);
                SendEmail(subject, body, filePath, to, cc);

                File.AppendAllText(
                    _logPath,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm} - ✅ {subject} sent to {string.Join(",", to)}{Environment.NewLine}"
                );
            }
            catch (Exception ex)
            {
                File.AppendAllText(
                  _logPath,
                  $"{DateTime.Now:yyyy-MM-dd HH:mm} ❌ Error generating '{subject}': {ex}{Environment.NewLine}"
                );
            }

        }

        private static void ExportQueryToExcel(string query, string filePath)
        {
            using var connection = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_dbPath};Persist Security Info=False;");
            using var command = new OleDbCommand(query, connection);
            var table = new DataTable();

            connection.Open();
            using var adapter = new OleDbDataAdapter(command);
            adapter.Fill(table);

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Results");

            // 1) Write headers:
            for (int c = 0; c < table.Columns.Count; c++)
                ws.Cell(1, c + 1).SetValue(table.Columns[c].ColumnName);

            var headerRow = ws.Row(1);
            headerRow.Style.Font.Bold = false;
            headerRow.Style.Fill.PatternType = XLFillPatternValues.Solid;
            headerRow.Style.Fill.BackgroundColor =
                XLColor.FromTheme(XLThemeColor.Background1, -0.25);
            ws.Columns(1, table.Columns.Count).AdjustToContents(1, 1);

            // 2) Write data rows:
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
            used.Style.Alignment.WrapText = true;
            used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            workbook.SaveAs(filePath);
        }



        private static void SendEmail(
            string subject,
            string bodyText,
            string attachmentPath,
            string[] to,
            string[] cc)
        {
            using var smtp = new SmtpClient("d250495a.ess.barracudanetworks.com", 587)
            {
                EnableSsl = false,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false
            };

            using var mail = new MailMessage
            {
                From = new MailAddress("no-reply@hagerman.com", "Hagerman & Company"),
                Subject = subject,
                Body = bodyText,
                IsBodyHtml = false
            };

            foreach (var addr in to)
                mail.To.Add(addr);
            foreach (var addr in cc)
                mail.CC.Add(addr);

            mail.Attachments.Add(new System.Net.Mail.Attachment(attachmentPath));

            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException("Attachment not found", attachmentPath);

            smtp.Send(mail);
        }

        private static void RunVerificationMacro()
        {
            var macroLogFile = Path.Combine(
                Path.GetDirectoryName(_logPath)!,
                "AccessMacroVerification.txt"
            );
            try
            {
                File.AppendAllText(
                    macroLogFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm} - 🔄 Running verification macro on database: {_dbPath}{Environment.NewLine}"
                );
                var accessApp = new Application();
                accessApp.OpenCurrentDatabase(_dbPath);
                accessApp.DoCmd.RunMacro("Run Process AFTER VERIFICATION");
                accessApp.CloseCurrentDatabase();
                accessApp.Quit();
                Marshal.ReleaseComObject(accessApp);

                File.AppendAllText(
                    macroLogFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm} - ✅ Verification macro completed successfully.{Environment.NewLine}"
                );
            }
            catch (Exception ex)
            {
                File.AppendAllText(
                    macroLogFile,
                    $"{DateTime.Now:yyyy-MM-dd HH:mm} - ❌ Error running verification macro: {ex.Message}{Environment.NewLine}"
                );
            }
        }

    }
}

using ClosedXML.Excel;
using Microsoft.Office.Interop.Access;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessMacroRunner.Services
{
    internal static class TwoQuarterAutomationService
    {
        private const string MacroName = "Run Autodesk Orders that do NOT Match NS SOs";
        private const string ReportQueryName = "Qry-Autodesk Orders without NS Matches";
        private const string EmailSubject = "Autodesk Orders with no Matching NS Orders";
        private const string EmailBody = "Here is a list of Autodesk Orders without NS Matches.";
        private const string PrimaryRecipient = "DanHarshman@hagerman.com";
        private static readonly string[] CcRecipients =
        {
            "ShellyAffeldt@hagerman.com",
            "ChrisDurham@hagerman.com"
        };

        private static readonly string DatabasePath = @"C:\ADSK-Automation\Autodesk Transaction Process.accdb";
        private static readonly string ImportFilePath = @"C:\ADSK-Automation\2Qtrs\JMHAutodeskSalesOrders2QtrsAllSyncResults.csv";
        private static readonly string OutputFolder = @"C:\ADSK-Automation\EmailReports";
        private static readonly string LogPath = @"C:\ADSK-Automation\Logs\TwoQuarterAutomationLog.txt";

        public static void Run()
        {
            DateTime runTime = DateTime.Now;
            EnsureDirectory(Path.GetDirectoryName(LogPath));
            EnsureDirectory(OutputFolder);

            Log("Starting two-quarter Autodesk order automation.");

            if (!File.Exists(DatabasePath))
                throw new FileNotFoundException("Access database not found.", DatabasePath);

            if (!File.Exists(ImportFilePath))
                throw new FileNotFoundException("Two-quarter NetSuite CSV not found.", ImportFilePath);

            int importedRows = ImportCsvToAccess(ImportFilePath);
            Log("Imported " + importedRows + " rows into SuiteTalk_Autodesk_Sales_Orders_2QTRS_ALL.");
            ArchiveSuccessfulInput(ImportFilePath, runTime);

            RunAccessMacro();

            string reportFolder = Path.Combine(OutputFolder, runTime.ToString("yyyy"), runTime.ToString("MM"));
            EnsureDirectory(reportFolder);
            string reportPath = Path.Combine(
                reportFolder,
                "Autodesk Orders Without NS Matches " + runTime.ToString("yyyy-MM-dd_HHmmss") + ".xlsx");

            ExportReport(reportPath);
            SendEmail(reportPath);

            Log("Two-quarter Autodesk order automation completed successfully.");
        }

        private static void ArchiveSuccessfulInput(string sourcePath, DateTime runTime)
        {
            try
            {
                string archiveFolder = Path.Combine(
                    Path.GetDirectoryName(sourcePath),
                    "Archive",
                    runTime.ToString("yyyy"),
                    runTime.ToString("MM"));
                EnsureDirectory(archiveFolder);

                string archivePath = Path.Combine(
                    archiveFolder,
                    Path.GetFileNameWithoutExtension(sourcePath) + "_" + runTime.ToString("yyyy-MM-dd_HH-mm-ss")
                    + Path.GetExtension(sourcePath));
                File.Copy(sourcePath, archivePath, false);
                Log("Archived successful two-quarter input to " + archivePath + ".");
            }
            catch (Exception ex)
            {
                // Archiving is valuable, but must not turn a successful import into a failed business run.
                Log("WARNING: Could not archive successful two-quarter input: " + ex.Message);
            }
        }

        private static int ImportCsvToAccess(string csvFilePath)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                + DatabasePath
                + ";Persist Security Info=False;";

            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var deleteCommand = new OleDbCommand(
                            "DELETE FROM [SuiteTalk_Autodesk_Sales_Orders_2QTRS_ALL]",
                            connection,
                            transaction))
                        {
                            deleteCommand.ExecuteNonQuery();
                        }

                        const string insertSql = @"
INSERT INTO [SuiteTalk_Autodesk_Sales_Orders_2QTRS_ALL]
    ([InternalID], [Customer], [SalesOrder], [AutodeskQuoteNbr],
     [LineID], [Autodesk_Line_], [Item], [Quantity],
     [UnitCustomerPrice], [AD_Start_Date], [AD_End_Date], [Order_Date])
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                        int importedRows = 0;
                        int lineNumber = 0;

                        using (var reader = new CsvRecordReader(csvFilePath))
                        using (var command = new OleDbCommand(insertSql, connection, transaction))
                        {
                            foreach (string[] values in reader.ReadRecords())
                            {
                                lineNumber++;

                                if (lineNumber == 1)
                                {
                                    ValidateHeader(values);
                                    continue;
                                }

                                if (values.Length != 13)
                                    throw new InvalidDataException(
                                        "CSV line " + lineNumber + " has " + values.Length + " columns; expected 13.");

                                command.Parameters.Clear();
                                command.Parameters.Add("@InternalID", OleDbType.VarWChar, 255).Value = RequiredText(values[0], lineNumber, "InternalID");
                                command.Parameters.Add("@Customer", OleDbType.VarWChar, 255).Value = DbText(values[1]);
                                command.Parameters.Add("@SalesOrder", OleDbType.VarWChar, 255).Value = DbText(values[2]);
                                command.Parameters.Add("@AutodeskQuoteNbr", OleDbType.VarWChar, 255).Value = DbText(values[3]);
                                command.Parameters.Add("@LineID", OleDbType.Double).Value = RequiredDouble(values[4], lineNumber, "Line ID");
                                command.Parameters.Add("@AutodeskLine", OleDbType.VarWChar, 255).Value = DbText(values[5]);
                                command.Parameters.Add("@Item", OleDbType.VarWChar, 255).Value = DbText(values[6]);
                                command.Parameters.Add("@Quantity", OleDbType.Double).Value = DbDouble(values[7], lineNumber, "Quantity");
                                command.Parameters.Add("@UnitCustomerPrice", OleDbType.Double).Value = DbDouble(values[9], lineNumber, "Unit Customer Price");
                                command.Parameters.Add("@StartDate", OleDbType.Date).Value = DbDate(values[10], lineNumber, "AD Start Date");
                                command.Parameters.Add("@EndDate", OleDbType.Date).Value = DbDate(values[11], lineNumber, "AD End Date");
                                command.Parameters.Add("@OrderDate", OleDbType.Date).Value = DbDate(values[12], lineNumber, "Order Date");

                                command.ExecuteNonQuery();
                                importedRows++;
                            }
                        }

                        if (lineNumber == 0)
                            throw new InvalidDataException("The two-quarter CSV is empty.");

                        transaction.Commit();
                        return importedRows;
                    }
                    catch
                    {
                        try { transaction.Rollback(); } catch { }
                        throw;
                    }
                }
            }
        }

        private static void ValidateHeader(string[] header)
        {
            string[] expected =
            {
                "InternalID", "Customer", "SO #", "Autodesk Quote #", "Line ID",
                "Autodesk Line #", "Item", "Quantity", "Item Rate", "Unit Customer Price",
                "AD Start Date", "AD End Date", "Order Date"
            };

            if (header.Length != expected.Length)
                throw new InvalidDataException(
                    "CSV header has " + header.Length + " columns; expected " + expected.Length + ".");

            for (int i = 0; i < expected.Length; i++)
            {
                string actual = (header[i] ?? string.Empty).Trim().TrimStart('\uFEFF');
                if (!actual.Equals(expected[i], StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        "CSV column " + (i + 1) + " is '" + actual + "'; expected '" + expected[i] + "'.");
                }
            }
        }

        private static string RequiredText(string value, int lineNumber, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException(fieldName + " is blank on CSV line " + lineNumber + ".");
            return value.Trim();
        }

        private static object DbText(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? (object)DBNull.Value : value.Trim();
        }

        private static double RequiredDouble(string value, int lineNumber, string fieldName)
        {
            object parsed = DbDouble(value, lineNumber, fieldName);
            if (parsed == DBNull.Value)
                throw new InvalidDataException(fieldName + " is blank on CSV line " + lineNumber + ".");
            return (double)parsed;
        }

        private static object DbDouble(string value, int lineNumber, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                return DBNull.Value;

            double number;
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out number) ||
                double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out number))
            {
                return number;
            }

            throw new InvalidDataException(
                fieldName + " has an invalid number on CSV line " + lineNumber + ": '" + value + "'.");
        }

        private static object DbDate(string value, int lineNumber, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                return DBNull.Value;

            DateTime date;
            string[] formats = { "M/d/yyyy", "MM/dd/yyyy", "M/d/yy", "MM/dd/yy" };
            if (DateTime.TryParseExact(value.Trim(), formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out date) ||
                DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
            {
                return date;
            }

            throw new InvalidDataException(
                fieldName + " has an invalid date on CSV line " + lineNumber + ": '" + value + "'.");
        }

        private static void RunAccessMacro()
        {
            Application accessApp = null;
            try
            {
                Log("Opening Access database and running macro '" + MacroName + "'.");
                accessApp = new Application();
                accessApp.OpenCurrentDatabase(DatabasePath, false);
                accessApp.DoCmd.RunMacro(MacroName);
                Log("Access macro completed.");
            }
            finally
            {
                if (accessApp != null)
                {
                    try { accessApp.CloseCurrentDatabase(); } catch { }
                    try { accessApp.Quit(); } catch { }
                    try { Marshal.FinalReleaseComObject(accessApp); } catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void ExportReport(string filePath)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                + DatabasePath
                + ";Persist Security Info=False;";

            var table = new DataTable();
            using (var connection = new OleDbConnection(connectionString))
            using (var command = new OleDbCommand("SELECT * FROM [" + ReportQueryName + "]", connection))
            using (var adapter = new OleDbDataAdapter(command))
            {
                connection.Open();
                adapter.Fill(table);
            }

            if (File.Exists(filePath))
                File.Delete(filePath);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Results");

                for (int column = 0; column < table.Columns.Count; column++)
                    worksheet.Cell(1, column + 1).SetValue(table.Columns[column].ColumnName);

                var headerRow = worksheet.Row(1);
                headerRow.Style.Font.Bold = false;
                headerRow.Style.Fill.PatternType = XLFillPatternValues.Solid;
                headerRow.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Background1, -0.25);
                worksheet.Columns(1, table.Columns.Count).AdjustToContents(1, 1);

                for (int row = 0; row < table.Rows.Count; row++)
                {
                    for (int column = 0; column < table.Columns.Count; column++)
                    {
                        object value = table.Rows[row][column];
                        worksheet.Cell(row + 2, column + 1).SetValue(value == DBNull.Value ? "" : value.ToString());
                    }
                }

                var used = worksheet.RangeUsed();
                if (used != null)
                {
                    used.Style.Alignment.WrapText = true;
                    used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                }

                workbook.SaveAs(filePath);
            }

            Log("Exported " + table.Rows.Count + " report rows to " + filePath + ".");
        }

        private static void SendEmail(string attachmentPath)
        {
            var settings = EmailSettings.Load();

            using (var smtp = new SmtpClient(settings.SmtpHost, settings.SmtpPort)
            {
                EnableSsl = settings.EnableSsl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false
            })
            using (var message = new MailMessage())
            {
                message.From = new MailAddress(settings.FromAddress, settings.FromDisplayName);
                message.To.Add(PrimaryRecipient);
                foreach (string ccRecipient in CcRecipients)
                    message.CC.Add(ccRecipient);
                message.Subject = EmailSubject;
                message.Body = EmailBody;
                message.IsBodyHtml = false;
                message.Attachments.Add(new System.Net.Mail.Attachment(attachmentPath));
                smtp.Send(message);
            }

            Log(
                "Email sent to " + PrimaryRecipient
                + " (CC: " + string.Join(", ", CcRecipients) + ")"
                + " with attachment " + attachmentPath + ".");
        }

        private static void EnsureDirectory(string path)
        {
            if (!string.IsNullOrWhiteSpace(path) && !Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private static void Log(string message)
        {
            EnsureDirectory(Path.GetDirectoryName(LogPath));
            File.AppendAllText(
                LogPath,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " - " + message + Environment.NewLine);
        }

        private sealed class EmailSettings
        {
            public string SmtpHost { get; private set; }
            public int SmtpPort { get; private set; }
            public bool EnableSsl { get; private set; }
            public string FromAddress { get; private set; }
            public string FromDisplayName { get; private set; }

            public static EmailSettings Load()
            {
                return new EmailSettings
                {
                    SmtpHost = GetString("SmtpHost", "d250495a.ess.barracudanetworks.com"),
                    SmtpPort = GetInt("SmtpPort", 587),
                    EnableSsl = GetBool("SmtpEnableSsl", false),
                    FromAddress = GetString("MailFromAddress", "no-reply@hagerman.com"),
                    FromDisplayName = GetString("MailFromDisplayName", "Hagerman & Company")
                };
            }

            private static string GetString(string key, string fallback)
            {
                string value = ConfigurationManager.AppSettings[key];
                return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
            }

            private static int GetInt(string key, int fallback)
            {
                int value;
                return int.TryParse(ConfigurationManager.AppSettings[key], out value) ? value : fallback;
            }

            private static bool GetBool(string key, bool fallback)
            {
                bool value;
                return bool.TryParse(ConfigurationManager.AppSettings[key], out value) ? value : fallback;
            }
        }

        private sealed class CsvRecordReader : IDisposable
        {
            private readonly StreamReader _reader;

            public CsvRecordReader(string path)
            {
                _reader = new StreamReader(path, Encoding.UTF8, true);
            }

            public System.Collections.Generic.IEnumerable<string[]> ReadRecords()
            {
                var fields = new System.Collections.Generic.List<string>();
                var current = new StringBuilder();
                bool inQuotes = false;

                while (true)
                {
                    int next = _reader.Read();
                    if (next == -1)
                    {
                        if (inQuotes)
                            throw new InvalidDataException("CSV ended inside a quoted field.");

                        if (current.Length > 0 || fields.Count > 0)
                        {
                            fields.Add(current.ToString());
                            yield return fields.ToArray();
                        }
                        yield break;
                    }

                    char character = (char)next;
                    if (inQuotes)
                    {
                        if (character == '"')
                        {
                            if (_reader.Peek() == '"')
                            {
                                _reader.Read();
                                current.Append('"');
                            }
                            else
                            {
                                inQuotes = false;
                            }
                        }
                        else
                        {
                            current.Append(character);
                        }
                    }
                    else if (character == '"' && current.Length == 0)
                    {
                        inQuotes = true;
                    }
                    else if (character == ',')
                    {
                        fields.Add(current.ToString());
                        current.Clear();
                    }
                    else if (character == '\r' || character == '\n')
                    {
                        if (character == '\r' && _reader.Peek() == '\n')
                            _reader.Read();

                        fields.Add(current.ToString());
                        current.Clear();
                        yield return fields.ToArray();
                        fields.Clear();
                    }
                    else
                    {
                        current.Append(character);
                    }
                }
            }

            public void Dispose()
            {
                _reader.Dispose();
            }
        }
    }
}

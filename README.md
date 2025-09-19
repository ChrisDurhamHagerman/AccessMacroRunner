# AccessMacroRunner
AccessMacroRunner is a small Windows console app (Framework 4.8) that your automation calls to do “desktop-style” work you can’t do directly in the web API. It opens your Access database, runs the required macro(s), exports data, and (when requested) builds Excel reports and emails them. It ships with Excel libraries (ClosedXML/OpenXML) so it can generate .xlsx files without needing Excel installed.

How it’s used in the bigger flow

Your NetSuiteAutomation web app launches AccessMacroRunner.exe with simple command-line arguments (e.g., emailreports). The runner then:

  1.	Runs the Access macro to process data and export the table you need (e.g., the “Netsuite Import Data” CSV) for NetSuite re-import.
  
  2.	Optionally creates Excel workbooks from Access queries (e.g., Qry-NBE RA Orders, Qry-NBE Orders without Quotes) and emails them to stakeholders—this is handled by Services/EmailReportService.cs.
  All file paths, credentials, and email settings are kept in App.config.

What the pieces do

    •	Program.cs – Entry point; parses the argument (e.g., run macro/export vs. send email reports) and orchestrates steps.
    
    •	Services/EmailReportService.cs – Builds polished .xlsx reports with ClosedXML and sends them as email attachments (subjects, recipients, and filenames follow your conventions).
    
    •	App.config – Central config (paths to the Access DB, output folders, log locations, recipients, etc.).
    In short: NetSuiteAutomation handles web/API work; AccessMacroRunner handles local Access + Excel + email tasks on demand so the end-to-end pipeline completes automatically.


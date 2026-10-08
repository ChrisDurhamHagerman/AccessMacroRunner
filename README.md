# AccessMacroRunner

## Publish to the automation server

The `ServerRelease` publish profile builds the x64 Release configuration, backs up
the currently deployed executable, configuration, and DLL files, and copies the
new output to `\\ns-autodesk\c$\ADSK-Automation\Release`.

From a Visual Studio Developer PowerShell prompt, run:

```powershell
msbuild AccessMacroRunner.csproj /t:Publish /p:PublishProfile=ServerRelease
```

Because this is a classic .NET Framework project, Visual Studio's **Publish**
command opens the ClickOnce wizard. Do not use that wizard for the server
deployment; run the command above from a Visual Studio Developer PowerShell
prompt instead. The account running the command must have access to the server's
administrative share.

## Retention cleanup

Run retention as a separate scheduled command after the business automations:

```powershell
AccessMacroRunner.exe retention
```

Retention is deliberately shipped with `RetentionDryRun=true`. In this mode it
writes the actions it would take to
`C:\ADSK-Automation\Logs\Retention\RetentionLog_yyyy-MM.txt`, but it does not
move or delete files. Review at least one normal week of dry-run logs before
setting `RetentionDryRun=false` in the deployed configuration.

The configured policy is:

- Runner logs: 180 days, with monthly rotation of runner-owned active logs.
- NetSuite input files: 90 days.
- Generated NetSuite export history: 180 days.
- Archived raw Autodesk/Forma downloads: 30 days.
- Two-quarter inputs: 180 days.
- Email reports: 13 months.
- Failed-run artifacts: 180 days.

Expired files first move into `C:\ADSK-Automation\PendingDeletion`. They are
permanently deleted by a later retention run only after the 7-day quarantine
period. Cleanup errors are logged and do not stop other files from being
processed. The active integration filenames are not moved by retention.

New email reports, generated NetSuite export history, and two-quarter input
snapshots are stored beneath `yyyy\MM` folders to avoid overwriting prior runs.

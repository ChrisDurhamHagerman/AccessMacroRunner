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

# AccessMacroRunner

## Publish to the automation server

The `ServerRelease` publish profile builds the x64 Release configuration, backs up
the currently deployed executable, configuration, and DLL files, and copies the
new output to `\\ns-autodesk\c$\ADSK-Automation\Release`.

From a Visual Studio Developer PowerShell prompt, run:

```powershell
msbuild AccessMacroRunner.csproj /t:Publish /p:PublishProfile=ServerRelease
```

Visual Studio can also select `ServerRelease` from the Publish page. The account
running Visual Studio must have access to the server's administrative share.

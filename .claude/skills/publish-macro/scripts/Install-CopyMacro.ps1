<#
.SYNOPSIS
    Installs the CopyMacro macro ("Get Macro") in Tabular Editor (TE3 and TE2), pointing at this solution folder.

.DESCRIPTION
    The macro lists the macros in "TE Scripts\TE Scripts.cs" (the methods that are not public static), and copies the
    one you pick to the clipboard, ready to paste in a C# script tab and save as a macro. The code comes from
    Copy-Macro.ps1 (dotnet build + TE_Scripts.CopyMacroFromVSFile), so it is always the current CopyMacro logic and the
    same output publish-macro writes.

    The solution folder is written into the macro, so run this again if you move or re-clone the repository.
    The import is done by Import-Macros.ps1 from the te-macros plugin (TabularEditor-Scripts repo): it backs up
    MacroActions.json, refuses to write while Tabular Editor is running, and replaces a macro with the same name.

.EXAMPLE
    .\Install-CopyMacro.ps1 -WhatIf
    .\Install-CopyMacro.ps1
    .\Install-CopyMacro.ps1 -Name "Copy Macro" -MacrosFile "$env:LOCALAPPDATA\TabularEditor3\MacroActions.json"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    # Macro name in Tabular Editor. An existing macro with this name is replaced.
    [string] $Name = 'Get Macro',
    # Default: the repository holding this script
    [string] $SolutionFolder,
    # Import-Macros.ps1 of the te-macros plugin. Default: found in the TabularEditor-Scripts clone
    # ($env:TE_PUBLISH_REPO or next to this repo) or in the installed te-macros plugin.
    [string] $Importer,
    # Macros files to install into. Default: every installed Tabular Editor (TE3 and TE2).
    [string[]] $MacrosFile,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
# resolved here, not as the parameter default: Windows PowerShell 5.1 leaves $PSScriptRoot empty in param defaults with -File
if (-not $SolutionFolder) { $SolutionFolder = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path }

if (-not (Test-Path (Join-Path $SolutionFolder 'TE Scripts\TE Scripts.cs'))) { throw "Not the TE-Scripting-in-Visual-Studio folder: $SolutionFolder" }
$copyMacroScript = Join-Path $SolutionFolder '.claude\skills\publish-macro\scripts\Copy-Macro.ps1'
if (-not (Test-Path $copyMacroScript)) { throw "Copy-Macro.ps1 not found: $copyMacroScript" }

if (-not $Importer) {
    $relative = 'plugins\te-macros\skills\te-macro-skills\scripts\Import-Macros.ps1'
    $candidates = @()
    if ($env:TE_PUBLISH_REPO) { $candidates += Join-Path $env:TE_PUBLISH_REPO $relative }
    $candidates += Join-Path (Split-Path $SolutionFolder) "TabularEditor-Scripts\$relative"
    $candidates += @(Get-ChildItem "$env:USERPROFILE\.claude\plugins\cache\*\te-macros\*\skills\te-macro-skills\scripts\Import-Macros.ps1" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -ExpandProperty FullName)
    $Importer = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $Importer) { throw "Import-Macros.ps1 not found. Clone TabularEditor-Scripts next to this repo, set TE_PUBLISH_REPO, install the te-macros plugin, or pass -Importer." }
}

$template = @'
// MACRO: Name: __NAME__
// MACRO: ValidContexts: Model
// MACRO: Tooltip: Copies a macro from TE Scripts.cs (TE-Scripting-in-Visual-Studio) to the clipboard, with the helper classes it uses
#r "System.Windows.Forms"
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// Installed by Install-CopyMacro.ps1. Run it again if the solution folder moves.
string solutionFolder = @"__SOLUTION__";
string copyMacroScript = Path.Combine(solutionFolder, @".claude\skills\publish-macro\scripts\Copy-Macro.ps1");
string macroFile = Path.Combine(solutionFolder, @"TE Scripts\TE Scripts.cs");

#if TE3
ScriptHelper.WaitFormVisible = false;
#endif

if (!File.Exists(macroFile) || !File.Exists(copyMacroScript))
{
    Error(String.Format(@"TE-Scripting-in-Visual-Studio not found in '{0}'. Run Install-CopyMacro.ps1 again from the repository.", solutionFolder));
    return;
}

// macros are the class methods (not local functions, which are indented deeper) that are not public static (same rule as CopyMacro)
var macroNames = new List<string>();
foreach (Match m in Regex.Matches(File.ReadAllText(macroFile), @"^(?: {8}|\t{2})((?:(?:public|private|protected|internal|static)\s+)*)void\s+(\w+)\s*\(", RegexOptions.Multiline))
{
    if (Regex.Replace(m.Groups[1].Value.Trim(), @"\s+", " ") != "public static") macroNames.Add(m.Groups[2].Value);
}

var form = new Form { Text = "Choose a macro", Width = 500, Height = 450, StartPosition = FormStartPosition.CenterScreen, Padding = new Padding(10) };
var filter = new TextBox { Dock = DockStyle.Top };
var list = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 40, FlowDirection = FlowDirection.RightToLeft };
var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
var ok = new Button { Text = "Copy", DialogResult = DialogResult.OK };
buttons.Controls.Add(cancel);
buttons.Controls.Add(ok);
form.Controls.Add(list);
form.Controls.Add(filter);
form.Controls.Add(buttons);
form.AcceptButton = ok;
form.CancelButton = cancel;
Action refresh = () =>
{
    list.Items.Clear();
    foreach (string n in macroNames.Where(n => n.IndexOf(filter.Text, StringComparison.OrdinalIgnoreCase) >= 0)) list.Items.Add(n);
    if (list.Items.Count > 0) list.SelectedIndex = 0;
};
filter.TextChanged += (s, e) => refresh();
list.DoubleClick += (s, e) => { form.DialogResult = DialogResult.OK; form.Close(); };
refresh();
if (form.ShowDialog() != DialogResult.OK || list.SelectedItem == null) return;
string macroName = list.SelectedItem.ToString();

// build the solution and run CopyMacro for that method in a separate process (the dll is never locked by Tabular Editor)
string outFile = Path.Combine(Path.GetTempPath(), "te-copy-macro-" + macroName + ".cs");
var psi = new ProcessStartInfo("powershell.exe",
    String.Format(@"-NoProfile -ExecutionPolicy Bypass -File ""{0}"" -Method {1} -OutFile ""{2}""", copyMacroScript, macroName, outFile))
{
    UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true
};
string stdOut, stdErr;
int exitCode;
using (var p = Process.Start(psi))
{
    var errTask = p.StandardError.ReadToEndAsync();
    stdOut = p.StandardOutput.ReadToEnd();
    p.WaitForExit();
    stdErr = errTask.Result;
    exitCode = p.ExitCode;
}
if (exitCode != 0 || !File.Exists(outFile))
{
    Error(String.Format("CopyMacro failed for {0}:\n\n{1}\n{2}", macroName, stdErr, stdOut));
    return;
}
Clipboard.SetText(File.ReadAllText(outFile));
File.Delete(outFile);
Info(String.Format("{0} copied to the clipboard. Paste it in a new C# script tab (CTRL+V) and save it as a macro.", macroName));
'@

$code = $template.Replace('__NAME__', $Name).Replace('__SOLUTION__', $SolutionFolder.TrimEnd('\'))
$folder = Join-Path ([IO.Path]::GetTempPath()) 'te-install-copymacro'
New-Item -ItemType Directory -Force $folder -WhatIf:$false | Out-Null
$file = Join-Path $folder (($Name -replace '[\\/:*?"<>|]', '-') + '.csx')
[IO.File]::WriteAllText($file, $code, (New-Object Text.UTF8Encoding($false)))

$importArgs = @{ Source = $file; WhatIf = [bool]$WhatIfPreference }
if ($MacrosFile) { $importArgs.MacrosFile = $MacrosFile }
if ($Force) { $importArgs.Force = $true }
Write-Host "Solution folder: $SolutionFolder"
Write-Host "Importer:        $Importer"
& $Importer @importArgs

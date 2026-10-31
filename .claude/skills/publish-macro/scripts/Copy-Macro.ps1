<#
.SYNOPSIS
    Runs CopyMacro (TE_Scripts.CopyMacroFromVSFile) for one method of "TE Scripts\TE Scripts.cs" and writes the macro code to a file.

.DESCRIPTION
    Same output as the CopyMacro macro in Tabular Editor, without the macro picker and the clipboard.
    Builds the "TE Scripts" project first (skip with -NoBuild) so the code reflects the current sources.

.EXAMPLE
    .\Copy-Macro.ps1 -Method SUM -OutFile "C:\...\TabularEditor-Scripts\Advanced\Skill-Ready Macros\Sum of Columns.cs"
#>
param(
    [Parameter(Mandatory)] [string] $Method,
    [Parameter(Mandatory)] [string] $OutFile,
    [string] $SolutionFolder = (Resolve-Path "$PSScriptRoot\..\..\..\..").Path,
    [switch] $NoBuild
)

$ErrorActionPreference = 'Stop'

$project = Join-Path $SolutionFolder 'TE Scripts\TE Scripts.csproj'
$binFolder = Join-Path $SolutionFolder 'TE Scripts\bin\Debug\net48'

if (-not $NoBuild) {
    $buildOutput = & dotnet build $project -c Debug -nologo -v q 2>&1
    if ($LASTEXITCODE -ne 0) {
        $buildOutput | Write-Host
        throw "Build of '$project' failed."
    }
}

# resolve dependencies (Roslyn, System.Collections.Immutable...) from the build output, whatever version is asked for
$resolver = [System.ResolveEventHandler] {
    param($sender, $e)
    $candidate = Join-Path $binFolder ((New-Object System.Reflection.AssemblyName $e.Name).Name + '.dll')
    if (Test-Path $candidate) { return [System.Reflection.Assembly]::LoadFrom($candidate) }
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)

try {
    $assembly = [System.Reflection.Assembly]::LoadFrom((Join-Path $binFolder 'TE Scripts.dll'))
    $type = $assembly.GetType('TE_Scripting.TE_Scripts', $true)
    $copyMacro = $type.GetMethod('CopyMacroFromVSFile')

    $outFolder = Split-Path $OutFile -Parent
    if ($outFolder -and -not (Test-Path $outFolder)) { New-Item -ItemType Directory -Path $outFolder | Out-Null }
    $OutFile = [System.IO.Path]::GetFullPath($OutFile)

    [string[]] $arguments = @(
        (Join-Path $SolutionFolder 'TE Scripts\TE Scripts.cs'),
        (Join-Path $SolutionFolder 'GeneralFunctions\GeneralFunctions.cs'),
        (Join-Path $SolutionFolder 'Report\Report.cs'),
        (Join-Path $SolutionFolder 'ReportFunctions\ReportFunctions.cs'),
        (Join-Path $SolutionFolder 'DaxUserDefinedFunction\DaxUserDefinedFunction.cs'),
        $Method,
        $OutFile
    )
    try {
        $copyMacro.Invoke($null, [object[]]$arguments) | Out-Null
    }
    catch [System.Reflection.TargetInvocationException] {
        throw $_.Exception.InnerException
    }
    Write-Output $OutFile
}
finally {
    [System.AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)
}

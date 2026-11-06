<#
.SYNOPSIS
Advanced command-line entry point (most people should double-click "Export Figures.bat").

.EXAMPLE
powershell -NoProfile -ExecutionPolicy Bypass -File .\engine\Run-Export.ps1 -Source .\MyFigures.pptx -FullWidthMm 165.1

Uses the private Python environment created by "Check Setup.bat" / "Export Figures.bat"
(%LOCALAPPDATA%\PaperFigureExporter\venv). Falls back to python/py on PATH, which then
needs: python -m pip install -r engine\requirements.txt
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true,Position=0)][string]$Source,
    [string]$Output,
    [string]$Config,
    [double]$FullWidthMm = 0,
    [double]$HalfWidthMm = 0,
    [switch]$AnalyzeOnly,
    [double]$MinPpi = 600,
    [int]$PreviewDpi = 600,
    [switch]$StrictPpi,
    [switch]$NativeImages,
    [switch]$IncludeHidden,
    [ValidateSet('review','clip','expand')][string]$Offslide = 'clip',
    [switch]$KeepGuideLines,
    [ValidateSet('match-full','fill')][string]$HalfWidthScale = 'match-full',
    [double]$MaxFileMb = 0
)
$ErrorActionPreference = 'Stop'
$env:PYTHONUTF8 = '1'
$env:PYTHONDONTWRITEBYTECODE = '1'
# Ignore other Python installs' search paths for this window only.
foreach ($name in 'PYTHONPATH', 'PYTHONHOME', 'PYTHONSTARTUP') { Remove-Item -LiteralPath "Env:$name" -ErrorAction SilentlyContinue }
$pythonPath = $null
$prefixArgs = @()
$privatePython = Join-Path $env:LOCALAPPDATA 'PaperFigureExporter\venv\Scripts\python.exe'
if (Test-Path -LiteralPath $privatePython) { $pythonPath = $privatePython }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $pythonPath = (Get-Command python).Source }
elseif (Get-Command py -ErrorAction SilentlyContinue) { $pythonPath = (Get-Command py).Source; $prefixArgs = @('-3') }
else { throw 'Python 3.10+ is required. Double-click "Check Setup.bat" to install everything automatically.' }
$scriptPath = Join-Path $PSScriptRoot 'export_figures.py'
$invokeArgs = $prefixArgs + @($scriptPath, $Source)
# Keep the renderer in the same PowerShell edition used to launch this script.
# In particular, powershell.exe must not silently select pwsh from PATH.
$rendererShell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
if ([IO.Path]::GetFileName($rendererShell) -notin @('powershell.exe','pwsh.exe','pwsh')) {
    $rendererShell = Join-Path $PSHOME 'powershell.exe'
}
$invokeArgs += @('--powershell-exe', $rendererShell)
if ($Output) { $invokeArgs += @('--output', $Output) }
if ($Config) { $invokeArgs += @('--config', $Config) }
if ($FullWidthMm -gt 0) { $invokeArgs += @('--full-width-mm', $FullWidthMm.ToString([Globalization.CultureInfo]::InvariantCulture)) }
if ($HalfWidthMm -gt 0) { $invokeArgs += @('--half-width-mm', $HalfWidthMm.ToString([Globalization.CultureInfo]::InvariantCulture)) }
if ($AnalyzeOnly) { $invokeArgs += '--analyze-only' }
$invokeArgs += @('--min-ppi', $MinPpi.ToString([Globalization.CultureInfo]::InvariantCulture), '--preview-dpi', $PreviewDpi.ToString([Globalization.CultureInfo]::InvariantCulture))
if ($StrictPpi) { $invokeArgs += '--strict-ppi' }
if ($NativeImages) { $invokeArgs += '--native-images' }
if ($IncludeHidden) { $invokeArgs += '--include-hidden' }
$invokeArgs += @('--offslide', $Offslide)
if ($KeepGuideLines) { $invokeArgs += '--keep-guide-lines' }
$invokeArgs += @('--half-width-scale', $HalfWidthScale)
if ($MaxFileMb -gt 0) { $invokeArgs += @('--max-file-mb', $MaxFileMb.ToString([Globalization.CultureInfo]::InvariantCulture)) }
& $pythonPath @invokeArgs
exit $LASTEXITCODE

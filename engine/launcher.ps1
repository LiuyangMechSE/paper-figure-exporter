<#
.SYNOPSIS
One-click launcher for Paper Figure Exporter (used by "Export Figures.bat").

.DESCRIPTION
1. Checks the computer: Windows, PowerShell, Microsoft PowerPoint, Python.
2. Installs what is missing:
     - Python 3.12 for the current user (winget, or the signed python.org
       installer as a fallback). No administrator rights are needed.
     - A private Python environment with the required packages, stored in
       %LOCALAPPDATA%\PaperFigureExporter (never inside this folder, so the
       tool folder stays small and safe to share or sync).
3. Asks for a PowerPoint file (or uses the files dropped onto the .bat).
4. Exports every figure to its own PDF + PNG preview and opens the results.

Parameters are optional; the .bat passes dropped file paths through.
This file is plain ASCII on purpose (Windows PowerShell 5.1 compatibility).
#>
[CmdletBinding()]
param(
    [Parameter(Position=0, ValueFromRemainingArguments=$true)][string[]]$Source,
    [switch]$CheckOnly,
    [switch]$Repair,
    [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2

# ---------------------------------------------------------------- constants
$EngineDir     = $PSScriptRoot
$AppDir        = Split-Path -Parent $EngineDir
$DataDir       = Join-Path $env:LOCALAPPDATA 'PaperFigureExporter'
$VenvDir       = Join-Path $DataDir 'venv'
$VenvPython    = Join-Path $VenvDir 'Scripts\python.exe'
$StampFile     = Join-Path $VenvDir 'requirements.sha256'
$LogDir        = Join-Path $DataDir 'logs'
$Requirements  = Join-Path $EngineDir 'requirements.txt'
$Exporter      = Join-Path $EngineDir 'export_figures.py'
$SettingsPath  = Join-Path $AppDir 'settings.json'
$MinPython     = [Version]'3.10'
$MaxPreferred  = [Version]'3.14'       # newest version with known-good wheels
$InstallVer    = '3.12.10'             # python.org fallback installer
$WingetId      = 'Python.Python.3.12'
$CheckScript   = Join-Path $EngineDir 'check_env.py'
$PipLog        = Join-Path $LogDir 'pip.log'
$VcRedistUrl   = 'https://aka.ms/vs/17/release/vc_redist.x64.exe'
$DeckFilter    = 'PowerPoint presentations (*.pptx;*.pptm;*.ppsx;*.ppsm)|*.pptx;*.pptm;*.ppsx;*.ppsm|All files (*.*)|*.*'

# Make Python and the console agree on UTF-8 (non-English paths and labels).
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONDONTWRITEBYTECODE = '1'     # keep the tool folder free of __pycache__
$env:PIP_DISABLE_PIP_VERSION_CHECK = '1'
$env:PYTHONNOUSERSITE = '1'
# Settings left behind by other Python installs (Anaconda, old projects) can make
# the private environment load foreign packages or install packages elsewhere.
# Clear them for this window only; nothing on the computer is changed.
foreach ($name in 'PYTHONPATH', 'PYTHONHOME', 'PYTHONSTARTUP', 'PIP_TARGET', 'PIP_PREFIX', 'PIP_USER', 'PIP_ROOT', 'PIP_REQUIRE_VIRTUALENV') {
    Remove-Item -LiteralPath "Env:$name" -ErrorAction SilentlyContinue
}
# Let Windows networking (used for downloads below) pass a signed-in company proxy.
try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }
try { [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false) } catch { }
try { $Host.UI.RawUI.WindowTitle = 'Paper Figure Exporter' } catch { }

# ---------------------------------------------------------------- helpers
function Say([string]$Text, [string]$Color = 'Gray') { Write-Host $Text -ForegroundColor $Color }
function Ok([string]$Text)   { Write-Host '  [ OK ] ' -ForegroundColor Green -NoNewline;  Write-Host $Text }
function Warn([string]$Text) { Write-Host '  [ !! ] ' -ForegroundColor Yellow -NoNewline; Write-Host $Text }
function Bad([string]$Text)  { Write-Host '  [FAIL] ' -ForegroundColor Red -NoNewline;    Write-Host $Text }
function Step([string]$Text) { Write-Host ''; Write-Host $Text -ForegroundColor Cyan }

function Finish([int]$Code) {
    try { Stop-Transcript | Out-Null } catch { }
    if (-not $NoPause) {
        Write-Host ''
        [void](Read-Host 'Press Enter to close this window')
    }
    exit $Code
}

function Ask-YesNo([string]$Question) {
    if ($NoPause) { return $true }
    $answer = Read-Host "$Question [Y/n]"
    return ($answer -eq '' -or $answer -match '^(y|yes)$')
}

# Run a native program, stream its output (stdout + stderr) to the console,
# and return the exit code. Native stderr must not become a terminating error.
function Invoke-Native([string]$Exe, [string[]]$Arguments) {
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Exe @Arguments 2>&1 | ForEach-Object { Write-Host ([string]$_) }
        return [int]$LASTEXITCODE
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        return 9999
    } finally { $ErrorActionPreference = $previous }
}

# Run a native program quietly and return its text output (or $null).
function Get-NativeOutput([string]$Exe, [string[]]$Arguments) {
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $text = (& $Exe @Arguments 2>$null | Out-String)
        if ($LASTEXITCODE -ne 0) { return $null }
        return $text
    } catch { return $null } finally { $ErrorActionPreference = $previous }
}

function Get-FileSha256([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-', '') }
    finally { $hasher.Dispose(); $stream.Dispose() }
}

# ---------------------------------------------------------------- Python discovery
$ProbeCode = "import sys, struct; print('PYOK|%d.%d.%d|%d|%s' % (sys.version_info[0], sys.version_info[1], sys.version_info[2], struct.calcsize('P') * 8, sys.executable))"

function Test-PythonCandidate([string]$Exe, [string[]]$Prefix = @()) {
    if (-not $Exe) { return $null }
    $out = Get-NativeOutput $Exe ($Prefix + @('-c', $ProbeCode))
    if (-not $out) { return $null }
    $line = @($out -split "`r?`n" | Where-Object { $_ -like 'PYOK|*' }) | Select-Object -First 1
    if (-not $line) { return $null }
    $parts = $line.Split('|')
    try { $version = [Version]$parts[1] } catch { return $null }
    if (-not (Test-Path -LiteralPath $parts[3])) { return $null }
    return [pscustomobject]@{ Path = $parts[3]; Version = $version; Bits = [int]$parts[2] }
}

function Find-BasePython {
    $found = @()
    # Python launcher first: it knows every registered installation.
    $launchers = @()
    $launchers += @(Get-Command py -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Source })
    $launchers += @("$env:LOCALAPPDATA\Programs\Python\Launcher\py.exe", "$env:windir\py.exe")
    foreach ($launcher in ($launchers | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
        $listing = Get-NativeOutput $launcher @('--list-paths')
        if ($listing) {
            foreach ($match in [regex]::Matches($listing, '([A-Za-z]:\\[^\r\n*]*?python\.exe)')) {
                $candidate = Test-PythonCandidate ($match.Groups[1].Value.Trim())
                if ($candidate) { $found += $candidate }
            }
        }
        $candidate = Test-PythonCandidate $launcher @('-3')
        if ($candidate) { $found += $candidate }
    }
    # Then anything on PATH and the standard install folders.
    $paths = @()
    foreach ($name in 'python', 'python3') {
        $paths += @(Get-Command $name -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Source })
    }
    foreach ($pattern in @("$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe",
                           "$env:ProgramFiles\Python3*\python.exe",
                           "$env:SystemDrive\Python3*\python.exe")) {
        $paths += @(Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    }
    foreach ($path in ($paths | Where-Object { $_ } | Select-Object -Unique)) {
        $candidate = Test-PythonCandidate $path
        if ($candidate) { $found += $candidate }
    }
    $usable = @($found | Where-Object { $_.Version -ge $MinPython } |
                Sort-Object -Property Path -Unique)
    if ($usable.Count -eq 0) { return $null }
    # Prefer 64-bit, then a version with known-good package wheels, then newest.
    return $usable | Sort-Object -Property `
        @{ Expression = { $_.Bits -eq 64 }; Descending = $true },
        @{ Expression = { [Version]('{0}.{1}' -f $_.Version.Major, $_.Version.Minor) -le $MaxPreferred }; Descending = $true },
        @{ Expression = { $_.Version }; Descending = $true } | Select-Object -First 1
}

function Install-Python {
    Say "  Python $MinPython or newer was not found on this computer."
    if (-not (Ask-YesNo "  Install Python 3.12 for your user account now? (free, no admin rights needed)")) {
        return $false
    }
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Say '  Installing Python with winget (this can take a few minutes)...'
        $code = Invoke-Native $winget.Source @('install', '--id', $WingetId, '--exact', '--scope', 'user',
                   '--silent', '--accept-package-agreements', '--accept-source-agreements')
        if (Find-BasePython) { return $true }
        Warn "winget did not finish the install (code $code). Trying the python.org installer instead."
    }
    $arch = 'amd64'
    $url = "https://www.python.org/ftp/python/$InstallVer/python-$InstallVer-$arch.exe"
    $installer = Join-Path $env:TEMP "python-$InstallVer-$arch.exe"
    try {
        Say "  Downloading $url"
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $oldProgress = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
        try { Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing } finally { $ProgressPreference = $oldProgress }
        $signature = Get-AuthenticodeSignature -FilePath $installer
        if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch 'Python Software Foundation') {
            throw "The downloaded installer is not validly signed by the Python Software Foundation ($($signature.Status))."
        }
        Say '  Running the Python installer silently...'
        $process = Start-Process -FilePath $installer -Wait -PassThru -ArgumentList @(
            '/quiet', 'InstallAllUsers=0', 'InstallLauncherAllUsers=0', 'PrependPath=1',
            'Include_launcher=1', 'Include_test=0', 'Include_doc=0', 'Shortcuts=0')
        if ($process.ExitCode -notin 0, 3010) { throw "Python installer exit code $($process.ExitCode)." }
    } catch {
        Bad "Automatic Python install failed: $($_.Exception.Message)"
        return $false
    } finally {
        Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
    }
    return [bool](Find-BasePython)
}

# ---------------------------------------------------------------- private environment
function New-PrivateEnvironment {
    Remove-Item -LiteralPath $VenvDir -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $VenvDir) {
        Bad "Could not delete the old environment folder (is another export window open?): $VenvDir"
        return $null
    }
    $base = Find-BasePython
    if (-not $base) {
        if (-not (Install-Python)) {
            Bad 'Python is required. Install it from https://www.python.org/downloads/ (tick "Add python.exe to PATH"), then run this again.'
            return $null
        }
        $base = Find-BasePython
    }
    Ok "Python found: $($base.Version) ($($base.Bits)-bit) at $($base.Path)"
    Say '  Creating the private Python environment...'
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
    $code = Invoke-Native $base.Path @('-m', 'venv', $VenvDir)
    if ($code -ne 0 -or -not (Test-Path -LiteralPath $VenvPython)) {
        Bad "Could not create a Python environment with $($base.Path) (code $code)."
        return $null
    }
    $venv = Test-PythonCandidate $VenvPython
    if (-not $venv) { Bad "The new Python environment does not start: $VenvPython"; return $null }
    return $venv
}

# ---------------------------------------------------------------- packages
# Runs check_env.py: imports each package and reports which one fails and why.
function Test-Packages {
    $lines = @()
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $lines = @(& $VenvPython $CheckScript 2>&1 | ForEach-Object { [string]$_ }) }
    catch { $lines = @([string]$_.Exception.Message) }
    finally { $ErrorActionPreference = $previous }
    $ok = $false
    $packages = @()
    $json = @($lines | Where-Object { $_ -like 'ENVCHECK *' }) | Select-Object -Last 1
    if ($json) {
        try {
            $data = $json.Substring(9) | ConvertFrom-Json
            $packages = @($data.packages)
            $ok = [bool]$data.ok
        } catch { }
    }
    return [pscustomobject]@{ Ok = $ok; Packages = $packages; Output = $lines }
}

function Show-PackageProblems($Check) {
    $bad = @($Check.Packages | Where-Object { -not $_.ok })
    if ($bad.Count -gt 0) {
        foreach ($p in $bad) { Say ("         {0}: {1}" -f $p.package, $p.error) 'Yellow' }
    } else {
        foreach ($line in @($Check.Output | Select-Object -Last 15)) { Say "         $line" 'Yellow' }
    }
}

function Get-ProblemText($Check) {
    return ((@($Check.Packages | ForEach-Object { [string]$_.error }) + @($Check.Output)) -join "`n")
}

# True when every package is installed but at least one still fails to load
# for a reason a clean environment can fix (not a missing Windows component).
function Test-RebuildHelps($Check) {
    if (@($Check.Packages).Count -eq 0) { return $true }
    if (@($Check.Packages | Where-Object { -not $_.version }).Count -gt 0) { return $false }
    return -not ((Get-ProblemText $Check) -match 'DLL load failed|VCRUNTIME|MSVCP')
}

function Get-SystemProxy {
    try {
        $target = [Uri]'https://pypi.org/simple/'
        $proxy = [Net.WebRequest]::GetSystemWebProxy().GetProxy($target)
        if ($proxy -and $proxy.Host -ne $target.Host) { return $proxy.GetLeftPart([UriPartial]::Authority) }
    } catch { }
    return $null
}

function Test-PyPI {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        [void](Invoke-WebRequest -Uri 'https://pypi.org/simple/pip/' -Method Head -UseBasicParsing -TimeoutSec 20)
        return $true
    } catch { return $false }
}

# Current pip versions trust the Windows certificate store, which networks that
# inspect HTTPS (company/university firewalls, antivirus web shields) require.
# Older pip cannot download itself there, so fetch the newest pip with Windows
# networking (which already trusts that store), verify its checksum, install it.
function Update-PipViaWindows {
    $wheelPath = $null
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $info = Invoke-RestMethod -Uri 'https://pypi.org/pypi/pip/json' -UseBasicParsing -TimeoutSec 60
        $wheel = @($info.urls | Where-Object { $_.packagetype -eq 'bdist_wheel' }) | Select-Object -First 1
        if (-not $wheel) { throw 'PyPI lists no pip wheel.' }
        $wheelPath = Join-Path $env:TEMP ([string]$wheel.filename)
        $oldProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try { Invoke-WebRequest -Uri $wheel.url -OutFile $wheelPath -UseBasicParsing -TimeoutSec 180 }
        finally { $ProgressPreference = $oldProgress }
        if ((Get-FileHash -LiteralPath $wheelPath -Algorithm SHA256).Hash -ne [string]$wheel.digests.sha256) {
            throw 'Checksum mismatch for the downloaded pip.'
        }
        Say "  Installing $($wheel.filename) (downloaded through Windows networking)" 'DarkGray'
        $code = Invoke-Native $VenvPython @('-m', 'pip', 'install', '--disable-pip-version-check', '--no-index', '--upgrade', $wheelPath)
        return ($code -eq 0)
    } catch {
        Warn "Could not update pip through Windows networking: $($_.Exception.Message)"
        return $false
    } finally {
        if ($wheelPath) { Remove-Item -LiteralPath $wheelPath -Force -ErrorAction SilentlyContinue }
    }
}

function Install-Packages {
    $common = @('-m', 'pip', 'install', '--disable-pip-version-check', '--progress-bar', 'off', '--log', $PipLog)
    Say '  Installing the required Python packages (first run takes 1-3 minutes)...'
    Say '  Step 1/2: updating the package installer (pip)' 'DarkGray'
    [void](Invoke-Native $VenvPython ($common + @('--upgrade', 'pip')))
    Say '  Step 2/2: downloading and installing packages' 'DarkGray'
    $code = Invoke-Native $VenvPython ($common + @('-r', $Requirements))
    $check = Test-Packages
    if ($code -eq 0 -and $check.Ok) { return $check }

    Warn "The first attempt did not work (pip exit code $code)."
    if (-not $check.Ok) { Show-PackageProblems $check }
    Say '  Retrying with network fixes (Windows certificates and proxy) and default pip settings...'
    [void](Update-PipViaWindows)
    # --isolated ignores pip.ini files and PIP_* variables that may point elsewhere.
    $retry = $common + @('--isolated', '--no-cache-dir')
    if ($code -eq 0) { $retry += '--force-reinstall' }
    $proxy = Get-SystemProxy
    if ($proxy) {
        Say "  Using the Windows proxy setting: $proxy"
        $retry += @('--proxy', $proxy)
    }
    [void](Invoke-Native $VenvPython ($retry + @('-r', $Requirements)))
    return (Test-Packages)
}

function Show-PackageFailure($Check) {
    Write-Host ''
    Bad 'The Python packages could not be set up. What went wrong:'
    Show-PackageProblems $Check
    Write-Host ''
    $problem = Get-ProblemText $Check
    $missing = @($Check.Packages | Where-Object { -not $_.version })
    if ($problem -match 'DLL load failed|VCRUNTIME|MSVCP') {
        Warn 'A Windows component (Microsoft Visual C++ Runtime) is missing or damaged.'
        Warn "Install it from $VcRedistUrl"
        Warn 'then restart the PC and run "Check Setup.bat" again.'
    } elseif ($missing.Count -gt 0 -and -not (Test-PyPI)) {
        Warn 'This computer cannot reach pypi.org, where Python packages are downloaded from.'
        Warn 'Check the internet connection, VPN or firewall (or try another network),'
        Warn 'then run "Check Setup.bat" again.'
    } elseif ($missing.Count -gt 0) {
        Warn 'pypi.org is reachable, but pip could not download from it. A firewall, proxy or'
        Warn 'antivirus "HTTPS scanning" may be blocking it; the messages above show the exact reason.'
    } else {
        Warn 'The packages are installed but do not load. The messages above show the exact reason.'
    }
    Say "  Details: $PipLog" 'DarkGray'
    Say "           $(Join-Path $LogDir 'last_run.log')" 'DarkGray'
}

# ---------------------------------------------------------------- environment check
function Initialize-Environment {
    $state = [ordered]@{ PowerPoint = $false; Python = $null }
    Step 'Checking this computer'

    if ($env:OS -ne 'Windows_NT') {
        Bad 'This tool runs on Windows (it drives Microsoft PowerPoint).'
        return $null
    }
    $osCaption = ''
    try { $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption } catch { $osCaption = [Environment]::OSVersion.VersionString }
    Ok "Windows: $osCaption"
    Ok "PowerShell: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"

    # Clear the downloaded-from-internet mark on our own scripts so the renderer
    # is never blocked. Only this tool's .ps1 files are touched.
    Get-ChildItem -LiteralPath $EngineDir -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        Unblock-File -ErrorAction SilentlyContinue

    $curVer = $null
    try { $curVer = (Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CurVer' -ErrorAction Stop).'(default)' } catch { }
    if (Test-Path -LiteralPath 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application\CLSID') {
        $state.PowerPoint = $true
        $label = if ($curVer) { $curVer } else { 'installed' }
        Ok "Microsoft PowerPoint: $label"
    } else {
        Warn 'Microsoft PowerPoint (desktop app) was not found.'
        Warn 'Figures cannot be exported without it; only the analysis report can be made.'
        Warn 'Install Microsoft 365 / Office desktop PowerPoint, then run this again.'
    }

    if (-not (Test-Path -LiteralPath $Requirements)) { Bad "Missing file: $Requirements"; return $null }
    if ($Repair -and (Test-Path -LiteralPath $VenvDir)) {
        Say '  Repair requested: removing the private Python environment...'
        Remove-Item -LiteralPath $VenvDir -Recurse -Force
    }

    # 1) Reuse the private environment when it still works, else build it.
    $venv = $null
    if (Test-Path -LiteralPath $VenvPython) {
        $venv = Test-PythonCandidate $VenvPython
        if (-not $venv -or $venv.Version -lt $MinPython) {
            Warn 'The private Python environment is broken (Python may have been removed). Rebuilding it.'
            $venv = $null
        }
    }
    if (-not $venv) { $venv = New-PrivateEnvironment }
    if (-not $venv) { return $null }
    Ok "Python environment: $($venv.Version) ($($venv.Bits)-bit)"

    # 2) Packages: install when requirements changed or an import fails.
    $wanted = Get-FileSha256 $Requirements
    $stamp = if (Test-Path -LiteralPath $StampFile) { (Get-Content -LiteralPath $StampFile -Raw).Trim() } else { '' }
    $check = Test-Packages
    if ($stamp -ne $wanted -or -not $check.Ok) {
        if ($stamp -eq $wanted) {
            Warn 'Some Python packages are missing or broken; reinstalling them:'
            Show-PackageProblems $check
        }
        $check = Install-Packages
        if (-not $check.Ok -and (Test-RebuildHelps $check)) {
            # Packages are present but still do not load: start from a clean environment once.
            Warn 'Rebuilding the private Python environment from scratch and trying once more...'
            $venv = New-PrivateEnvironment
            if (-not $venv) { return $null }
            $check = Install-Packages
        }
        if (-not $check.Ok) { Show-PackageFailure $check; return $null }
        Set-Content -LiteralPath $StampFile -Value $wanted -Encoding ASCII
    }
    foreach ($p in @($check.Packages | Where-Object { $_.outside_environment })) {
        Warn "$($p.package) is loaded from outside the private environment: $($p.location)"
    }
    Ok ('Packages: ' + ((@($check.Packages) | ForEach-Object { '{0} {1}' -f $_.package, $_.version }) -join ', '))
    $state.Python = $VenvPython
    return $state
}

# ---------------------------------------------------------------- settings
function Read-Settings {
    $settings = [ordered]@{
        full_width_mm = 165.1; half_width_mm = $null; min_ppi = 600; preview_dpi = 600
        include_hidden_slides = $false; strict_ppi = $false; open_results_when_done = $true
        offslide_content = 'clip'; half_width_text = 'match_full'; max_file_mb = 10
    }
    if (Test-Path -LiteralPath $SettingsPath) {
        try {
            $json = Get-Content -LiteralPath $SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($key in @($settings.Keys)) {
                if ($json.PSObject.Properties[$key]) { $settings[$key] = $json.$key }
            }
        } catch {
            Warn "settings.json could not be read ($($_.Exception.Message)); using defaults."
        }
    }
    return $settings
}

function Format-Number($Value) { return ([double]$Value).ToString([Globalization.CultureInfo]::InvariantCulture) }

# ---------------------------------------------------------------- file picking
function Select-Decks {
    Add-Type -AssemblyName System.Windows.Forms
    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Choose the PowerPoint file(s) that contain your figures'
    $dialog.Filter = $DeckFilter
    $dialog.Multiselect = $true
    try {
        if ($dialog.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { return @($dialog.FileNames) }
        return @()
    } finally { $dialog.Dispose(); $owner.Dispose() }
}

# ---------------------------------------------------------------- export one deck
function Export-Deck([string]$Deck, $State, $Settings) {
    $file = Get-Item -LiteralPath $Deck
    Step "Exporting figures from: $($file.Name)"
    $ext = $file.Extension.ToLowerInvariant()
    if ($ext -eq '.ppt' -or $ext -eq '.pps') {
        Bad 'Old .ppt format. Open it in PowerPoint and use File > Save As > PowerPoint Presentation (.pptx), then try again.'
        return 1
    }
    if ($ext -notin '.pptx', '.pptm', '.ppsx', '.ppsm') { Bad 'Not a PowerPoint file.'; return 1 }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $output = Join-Path $file.DirectoryName ("{0}_figures_{1}" -f $file.BaseName, $stamp)
    $arguments = @($Exporter, $file.FullName, '--output', $output,
                   '--min-ppi', (Format-Number $Settings.min_ppi),
                   '--preview-dpi', ([string][int]$Settings.preview_dpi))
    if ($Settings.full_width_mm) { $arguments += @('--full-width-mm', (Format-Number $Settings.full_width_mm)) }
    if ($Settings.half_width_mm) { $arguments += @('--half-width-mm', (Format-Number $Settings.half_width_mm)) }
    if ($Settings.include_hidden_slides) { $arguments += '--include-hidden' }
    # Parts of objects outside the slide are cut at the slide edge, as PowerPoint shows them.
    $offslide = ([string]$Settings.offslide_content).ToLowerInvariant()
    if ($offslide -notin 'clip', 'expand', 'review') {
        Warn "settings.json: offslide_content must be clip, expand or review (got '$offslide'); using clip."
        $offslide = 'clip'
    }
    $arguments += @('--offslide', $offslide)
    # Half-width figures: same printed text size as full-width ones (centred), or fill the half width.
    $halfMode = ([string]$Settings.half_width_text).ToLowerInvariant().Replace('_', '-')
    if ($halfMode -notin 'match-full', 'fill') {
        Warn "settings.json: half_width_text must be match_full or fill (got '$halfMode'); using match_full."
        $halfMode = 'match-full'
    }
    $arguments += @('--half-width-scale', $halfMode)
    if ($Settings.max_file_mb) { $arguments += @('--max-file-mb', (Format-Number $Settings.max_file_mb)) }
    if ($Settings.strict_ppi) { $arguments += '--strict-ppi' }

    # Saved grouping/width choices for this deck: "<deck name>.figures.json".
    $deckConfig = Join-Path $file.DirectoryName ($file.BaseName + '.figures.json')
    if (Test-Path -LiteralPath $deckConfig) {
        $configHash = $null
        try { $configHash = [string](Get-Content -LiteralPath $deckConfig -Raw -Encoding UTF8 | ConvertFrom-Json).source_sha256 } catch {
            Bad "$([IO.Path]::GetFileName($deckConfig)) is not valid JSON: $($_.Exception.Message)"
            return 1
        }
        if ($configHash -and $configHash -ne (Get-FileSha256 $file.FullName)) {
            Warn "$([IO.Path]::GetFileName($deckConfig)) was made for an older saved version of this deck."
            Warn 'Shape numbers may have changed, so it is ignored this time. Check the new review.html,'
            Warn 'then replace it with the new overrides.generated.json (see README.md).'
        } else {
            Say "  Using saved figure choices: $([IO.Path]::GetFileName($deckConfig))"
            $arguments += @('--config', $deckConfig)
        }
    }

    if ($State.PowerPoint) {
        # Keep the renderer in the same PowerShell edition as this launcher.
        $shell = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        if ([IO.Path]::GetFileName($shell) -notin 'powershell.exe', 'pwsh.exe') { $shell = Join-Path $PSHOME 'powershell.exe' }
        $arguments += @('--powershell-exe', $shell)
        Say '  PowerPoint will open in the background; please do not close it while this runs.'
    } else {
        $arguments += '--analyze-only'
        Warn 'PowerPoint not available: making the analysis report only (no PDFs).'
    }

    $code = Invoke-Native $State.Python $arguments
    Write-Host ''
    switch ($code) {
        0 { Ok 'Done. All figures were exported.' }
        2 { Warn 'Done, but some slides need your decision (shown in orange in review.html).'
            Warn 'See README.md > "When something needs review" for the 2-minute fix.' }
        3 { Warn 'Done, but some pictures are below the target resolution. See quality_report.html.' }
        default { Bad "The export stopped with an error (code $code). The messages above explain why." }
    }
    if (Test-Path -LiteralPath $output) {
        Say "  Results folder: $output" 'White'
        if ($Settings.open_results_when_done) {
            $review = Join-Path $output 'review.html'
            if (Test-Path -LiteralPath $review) { Invoke-Item -LiteralPath $review }
            Invoke-Item -LiteralPath $output
        }
    }
    return $code
}

# ---------------------------------------------------------------- main
try {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    try { Start-Transcript -LiteralPath (Join-Path $LogDir 'last_run.log') -Force | Out-Null } catch { }

    Say ''
    Say '  Paper Figure Exporter' 'White'
    Say '  PowerPoint slides -> one cropped vector PDF per figure' 'DarkGray'

    $state = Initialize-Environment
    if ($null -eq $state) { Finish 1 }
    if ($CheckOnly) {
        Step 'Setup check complete.'
        if ($state.PowerPoint) { Ok 'Ready. Double-click "Export Figures.bat" (or drop a .pptx onto it) to export.' }
        else { Warn 'Python is ready, but PowerPoint is missing, so only analysis reports are possible.' }
        Finish 0
    }

    $decks = @($Source | Where-Object { $_ -and $_.Trim() })
    if ($decks.Count -eq 0) {
        Step 'Choose your PowerPoint file in the window that just opened...'
        $decks = @(Select-Decks)
        if ($decks.Count -eq 0) { Say '  No file chosen. Nothing to do.'; Finish 0 }
    }
    $settings = Read-Settings
    $worst = 0
    foreach ($deck in $decks) {
        if (-not (Test-Path -LiteralPath $deck -PathType Leaf)) { Bad "File not found: $deck"; $worst = 1; continue }
        $code = Export-Deck $deck $state $settings
        if ($code -eq 1 -or $code -gt 3) { $worst = 1 } elseif ($code -gt $worst) { $worst = $code }
    }
    Finish $worst
} catch {
    Bad "Unexpected problem: $($_.Exception.Message)"
    Say "  Log file: $(Join-Path $LogDir 'last_run.log')" 'DarkGray'
    Finish 1
}

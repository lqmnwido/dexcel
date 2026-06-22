# install.ps1 - Dexcel installer for Windows
#
# Install:
#   powershell -ExecutionPolicy Bypass -c "irm https://dexcel.kohich.site/install.ps1 | iex"
#
# Run after install:
#   dexcel

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$ReleaseBaseUrl = $env:DEXCEL_RELEASE_URL
if (-not $ReleaseBaseUrl) {
    $ReleaseBaseUrl = "https://dexcel.kohich.site/releases/latest/download"
}

$InstallDir = Join-Path $env:USERPROFILE ".dexcel"
$AppDir = Join-Path $InstallDir "app"
$VenvDir = Join-Path $InstallDir "venv"
$LogDir = Join-Path $InstallDir "logs"
$BinDir = Join-Path $InstallDir "bin"
$LauncherPath = Join-Path $BinDir "dexcel.cmd"

$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$InstallLog = Join-Path $LogDir "install_$Timestamp.log"

New-Item -ItemType Directory -Force -Path $InstallDir, $AppDir, $LogDir, $BinDir | Out-Null

$script:ProgressWidth = 36

function Write-Log {
    param([string]$Message)

    Add-Content `
        -Path $InstallLog `
        -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message" `
        -Encoding UTF8
}

function Show-Header {
    Clear-Host

    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host " Dexcel Installer" -ForegroundColor Cyan
    Write-Host " Database to Excel exporter for Windows" -ForegroundColor DarkCyan
    Write-Host "====================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Install path : $InstallDir"
    Write-Host "Log file     : $InstallLog"
    Write-Host ""
}

function Show-Stage {
    param(
        [int]$Percent,
        [string]$Title
    )

    if ($Percent -lt 0) { $Percent = 0 }
    if ($Percent -gt 100) { $Percent = 100 }

    $Done = [math]::Floor(($Percent / 100) * $script:ProgressWidth)
    $Left = $script:ProgressWidth - $Done
    $Bar = ("#" * $Done) + ("-" * $Left)

    Write-Host ""
    Write-Host "----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[$Bar] $Percent%" -ForegroundColor Cyan
    Write-Host "$Title" -ForegroundColor White
    Write-Host "----------------------------------------------------" -ForegroundColor DarkGray

    Write-Log $Title
}

function Show-Task {
    param([string]$Message)

    Write-Host "  > $Message" -ForegroundColor Gray
    Write-Log $Message
}

function Show-Ok {
    param([string]$Message)

    Write-Host "  [OK] $Message" -ForegroundColor Green
    Write-Log "OK: $Message"
}

function Show-Warn {
    param([string]$Message)

    Write-Host "  [WARN] $Message" -ForegroundColor Yellow
    Write-Log "WARN: $Message"
}

function Show-FailLine {
    param([string]$Message)

    Write-Host "  [FAIL] $Message" -ForegroundColor Red
    Write-Log "FAIL: $Message"
}

function Fail {
    param([string]$Message)

    Write-Host ""
    Write-Host "====================================================" -ForegroundColor Red
    Write-Host " Installation failed" -ForegroundColor Red
    Write-Host "====================================================" -ForegroundColor Red
    Write-Host $Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Log file:"
    Write-Host "  $InstallLog" -ForegroundColor Yellow

    Write-Log "FATAL: $Message"
    exit 1
}

function Quote-Arg {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    if ($Value -match '[\s"]') {
        return '"' + ($Value -replace '"', '\"') + '"'
    }

    return $Value
}

function Join-Args {
    param([string[]]$Arguments)

    if (-not $Arguments) {
        return ""
    }

    return ($Arguments | ForEach-Object { Quote-Arg $_ }) -join " "
}

function Show-OutputTail {
    param(
        [string]$Text,
        [int]$Lines = 20
    )

    if (-not $Text) {
        return
    }

    $Parts = $Text -split "`r?`n"
    $Tail = $Parts | Where-Object { $_.Trim() -ne "" } | Select-Object -Last $Lines

    foreach ($Line in $Tail) {
        Write-Host "    $Line" -ForegroundColor DarkGray
    }
}

function Invoke-ProcessQuiet {
    param(
        [string]$Title,
        [string]$FilePath,
        [string[]]$Arguments,
        [switch]$AllowFailure,
        [switch]$Silent
    )

    if (-not $Arguments) {
        $Arguments = @()
    }

    Write-Log "RUN: $Title"
    Write-Log ("CMD: " + $FilePath + " " + (Join-Args $Arguments))

    $StartTime = Get-Date

    $Psi = New-Object System.Diagnostics.ProcessStartInfo
    $Psi.FileName = $FilePath
    $Psi.Arguments = Join-Args $Arguments
    $Psi.UseShellExecute = $false
    $Psi.RedirectStandardOutput = $true
    $Psi.RedirectStandardError = $true
    $Psi.CreateNoWindow = $true

    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $Psi

    try {
        [void]$Process.Start()

        $StdOutTask = $Process.StandardOutput.ReadToEndAsync()
        $StdErrTask = $Process.StandardError.ReadToEndAsync()

        $Spinner = @("|", "/", "-", "\")
        $SpinIndex = 0

        while (-not $Process.HasExited) {
            if (-not $Silent) {
                $Spin = $Spinner[$SpinIndex % $Spinner.Count]
                $SpinIndex++

                $Elapsed = [int]((Get-Date) - $StartTime).TotalSeconds
                $Line = "  [$Spin] $Title... ${Elapsed}s"

                Write-Host -NoNewline "`r$Line"
            }

            Start-Sleep -Milliseconds 180
        }

        $Process.WaitForExit()

        $StdOut = $StdOutTask.Result
        $StdErr = $StdErrTask.Result
        $Output = ($StdOut + "`n" + $StdErr).Trim()

        if ($Output) {
            Add-Content -Path $InstallLog -Value $Output -Encoding UTF8
        }

        $ElapsedTotal = [int]((Get-Date) - $StartTime).TotalSeconds
        $ExitCode = $Process.ExitCode

        if (-not $Silent) {
            $Clear = " " * 100
            Write-Host -NoNewline "`r$Clear`r"
        }

        if ($ExitCode -eq 0) {
            if (-not $Silent) {
                Show-Ok "$Title completed in ${ElapsedTotal}s"
            }

            return [pscustomobject]@{
                ExitCode = $ExitCode
                Output   = $Output
            }
        }

        if ($AllowFailure) {
            if (-not $Silent) {
                Show-Warn "$Title returned exit code $ExitCode"
            }

            return [pscustomobject]@{
                ExitCode = $ExitCode
                Output   = $Output
            }
        }

        if (-not $Silent) {
            Show-FailLine "$Title failed with exit code $ExitCode"

            if ($Output) {
                Write-Host ""
                Write-Host "  Last output:" -ForegroundColor Yellow
                Show-OutputTail $Output 25
            }
        }

        throw "$Title failed with exit code $ExitCode"
    } catch {
        if ($AllowFailure) {
            return [pscustomobject]@{
                ExitCode = 1
                Output   = $_.ToString()
            }
        }

        throw $_
    }
}

function Refresh-EnvironmentPath {
    $MachinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $UserPath = [Environment]::GetEnvironmentVariable("Path", "User")

    $env:Path = "$MachinePath;$UserPath"
}

function Test-PythonCommand {
    param([string]$Command)

    $Code = "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)"

    try {
        if ($Command -eq "py -3.12") {
            $Result = Invoke-ProcessQuiet "Test Python 3.12" "py" @("-3.12", "-c", $Code) -AllowFailure -Silent
            return $Result.ExitCode -eq 0
        }

        if ($Command -eq "py -3.11") {
            $Result = Invoke-ProcessQuiet "Test Python 3.11" "py" @("-3.11", "-c", $Code) -AllowFailure -Silent
            return $Result.ExitCode -eq 0
        }

        if ($Command -eq "py -3.10") {
            $Result = Invoke-ProcessQuiet "Test Python 3.10" "py" @("-3.10", "-c", $Code) -AllowFailure -Silent
            return $Result.ExitCode -eq 0
        }

        $Result = Invoke-ProcessQuiet "Test Python path" $Command @("-c", $Code) -AllowFailure -Silent
        return $Result.ExitCode -eq 0
    } catch {
        return $false
    }
}

function Find-Python {
    Refresh-EnvironmentPath

    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "py -3.12") { return "py -3.12" }
        if (Test-PythonCommand "py -3.11") { return "py -3.11" }
        if (Test-PythonCommand "py -3.10") { return "py -3.10" }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "python") { return "python" }
    }

    $CommonPaths = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
        "$env:ProgramFiles\Python312\python.exe",
        "$env:ProgramFiles\Python311\python.exe",
        "$env:ProgramFiles\Python310\python.exe"
    )

    foreach ($Path in $CommonPaths) {
        if (Test-Path $Path) {
            if (Test-PythonCommand $Path) {
                return $Path
            }
        }
    }

    return $null
}

function Get-PythonVersion {
    param([string]$PythonCmd)

    try {
        if ($PythonCmd -eq "py -3.12") {
            $Result = Invoke-ProcessQuiet "Get Python version" "py" @("-3.12", "--version") -AllowFailure -Silent
            return $Result.Output.Trim()
        }

        if ($PythonCmd -eq "py -3.11") {
            $Result = Invoke-ProcessQuiet "Get Python version" "py" @("-3.11", "--version") -AllowFailure -Silent
            return $Result.Output.Trim()
        }

        if ($PythonCmd -eq "py -3.10") {
            $Result = Invoke-ProcessQuiet "Get Python version" "py" @("-3.10", "--version") -AllowFailure -Silent
            return $Result.Output.Trim()
        }

        $Result = Invoke-ProcessQuiet "Get Python version" $PythonCmd @("--version") -AllowFailure -Silent
        return $Result.Output.Trim()
    } catch {
        return $PythonCmd
    }
}

function Install-Python312 {
    Show-Task "Compatible Python was not found"
    Show-Task "Installing Python 3.12"

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Show-Task "Trying winget install first"

        $WingetResult = Invoke-ProcessQuiet `
            "Install Python 3.12 using winget" `
            "winget" `
            @(
                "install",
                "-e",
                "--id",
                "Python.Python.3.12",
                "--silent",
                "--disable-interactivity",
                "--accept-package-agreements",
                "--accept-source-agreements"
            ) `
            -AllowFailure

        Refresh-EnvironmentPath

        $PythonAfterWinget = Find-Python
        if ($PythonAfterWinget) {
            Show-Ok "Python was installed successfully by winget"
            return $PythonAfterWinget
        }

        Show-Warn "winget finished but Python is not available in this terminal yet"
        Show-Warn "Falling back to direct Python installer"
    } else {
        Show-Warn "winget was not found"
    }

    $PythonInstaller = Join-Path $env:TEMP "python-3.12-installer.exe"
    $PythonInstallerUrl = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"

    try {
        Show-Task "Downloading Python 3.12 installer"

        Invoke-WebRequest `
            -Uri $PythonInstallerUrl `
            -OutFile $PythonInstaller `
            -UseBasicParsing

        $InstallResult = Invoke-ProcessQuiet `
            "Install Python 3.12 directly" `
            $PythonInstaller `
            @(
                "/quiet",
                "InstallAllUsers=0",
                "PrependPath=1",
                "Include_launcher=1",
                "Include_pip=1",
                "Include_test=0"
            ) `
            -AllowFailure

        Remove-Item $PythonInstaller -Force -ErrorAction SilentlyContinue

        Refresh-EnvironmentPath

        $PythonAfterDirect = Find-Python
        if ($PythonAfterDirect) {
            Show-Ok "Python was installed successfully"
            return $PythonAfterDirect
        }

        throw "Python installer completed, but Python could not be found in PATH or common install paths."
    } catch {
        Remove-Item $PythonInstaller -Force -ErrorAction SilentlyContinue
        throw $_
    }
}

function Create-Venv {
    param([string]$PythonCmd)

    if ($PythonCmd -eq "py -3.12") {
        Invoke-ProcessQuiet "Create Python virtual environment" "py" @("-3.12", "-m", "venv", $VenvDir)
        return
    }

    if ($PythonCmd -eq "py -3.11") {
        Invoke-ProcessQuiet "Create Python virtual environment" "py" @("-3.11", "-m", "venv", $VenvDir)
        return
    }

    if ($PythonCmd -eq "py -3.10") {
        Invoke-ProcessQuiet "Create Python virtual environment" "py" @("-3.10", "-m", "venv", $VenvDir)
        return
    }

    Invoke-ProcessQuiet "Create Python virtual environment" $PythonCmd @("-m", "venv", $VenvDir)
}

function Test-VenvCompatible {
    param([string]$PythonExe)

    if (-not (Test-Path $PythonExe)) {
        return $false
    }

    $Code = "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)"

    try {
        $Result = Invoke-ProcessQuiet "Test virtual environment Python" $PythonExe @("-c", $Code) -AllowFailure -Silent
        return $Result.ExitCode -eq 0
    } catch {
        return $false
    }
}

function Test-PythonImport {
    param(
        [string]$PythonExe,
        [string]$Module
    )

    $Code = "import importlib; importlib.import_module('$Module')"

    try {
        $Result = Invoke-ProcessQuiet "Test import $Module" $PythonExe @("-c", $Code) -AllowFailure -Silent

        return [pscustomobject]@{
            Success = ($Result.ExitCode -eq 0)
            Output  = $Result.Output
        }
    } catch {
        return [pscustomobject]@{
            Success = $false
            Output  = $_.ToString()
        }
    }
}

Show-Header

Show-Stage 10 "Checking compatible Python"

$PythonCmd = Find-Python

if (-not $PythonCmd) {
    try {
        $PythonCmd = Install-Python312
    } catch {
        Fail "Failed to install Python 3.12. $_"
    }
}

$PythonVersion = Get-PythonVersion $PythonCmd
Show-Ok "Using $PythonVersion via $PythonCmd"

Show-Stage 20 "Checking old virtual environment"

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (Test-Path $VenvPython) {
    if (Test-VenvCompatible $VenvPython) {
        Show-Ok "Existing virtual environment is compatible"
    } else {
        Show-Warn "Old virtual environment is incompatible"
        Show-Task "Removing old virtual environment"
        Remove-Item -Recurse -Force $VenvDir
        Show-Ok "Old virtual environment removed"
    }
} else {
    Show-Task "No existing virtual environment found"
}

Show-Stage 30 "Downloading Dexcel files"

try {
    $MainFile = Join-Path $AppDir "db_to_excel.py"
    $RequirementsFile = Join-Path $AppDir "requirements.txt"

    Show-Task "Downloading db_to_excel.py"
    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/db_to_excel.py" `
        -OutFile $MainFile `
        -UseBasicParsing

    Show-Task "Downloading requirements.txt"
    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/requirements.txt" `
        -OutFile $RequirementsFile `
        -UseBasicParsing

    Show-Ok "Application files downloaded"
} catch {
    Fail "Failed to download Dexcel files from $ReleaseBaseUrl. $_"
}

Show-Stage 40 "Preparing Python environment"

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (-not (Test-Path $VenvPython)) {
    try {
        Create-Venv $PythonCmd
    } catch {
        Fail "Failed to create virtual environment. $_"
    }

    if (-not (Test-Path $VenvPython)) {
        Fail "Virtual environment was created but python.exe was not found."
    }

    Show-Ok "Virtual environment ready"
} else {
    Show-Ok "Using existing virtual environment"
}

Show-Stage 55 "Upgrading installer tools"

try {
    Invoke-ProcessQuiet `
        "Upgrade pip, setuptools, and wheel" `
        $VenvPython `
        @("-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel")
} catch {
    Fail "Failed to upgrade pip, setuptools, and wheel. $_"
}

Show-Stage 70 "Installing Dexcel dependencies"

try {
    Invoke-ProcessQuiet `
        "Install Python dependencies" `
        $VenvPython `
        @("-m", "pip", "install", "-r", $RequirementsFile)
} catch {
    Write-Host ""
    Write-Host "Manual debug command:" -ForegroundColor Yellow
    Write-Host "  `"$VenvPython`" -m pip install -r `"$RequirementsFile`""

    Fail "Failed to install Dexcel dependencies. $_"
}

Show-Stage 80 "Verifying core packages"

try {
    Invoke-ProcessQuiet `
        "Verify pandas and openpyxl" `
        $VenvPython `
        @("-c", "import pandas, openpyxl")
} catch {
    Fail "Core packages failed verification. $_"
}

Show-Stage 88 "Checking database drivers"

$Drivers = [ordered]@{
    "MySQL/MariaDB" = "pymysql"
    "PostgreSQL"   = "psycopg"
    "SQL Server"   = "pyodbc"
    "Oracle"       = "oracledb"
}

$Working = New-Object System.Collections.Generic.List[string]
$Broken = New-Object System.Collections.Generic.List[string]

foreach ($Label in $Drivers.Keys) {
    $Module = $Drivers[$Label]

    $ImportResult = Test-PythonImport -PythonExe $VenvPython -Module $Module

    if ($ImportResult.Success) {
        $Working.Add($Label)
        Show-Ok "$Label driver available"
        Write-Log "Driver OK: $Label ($Module)"
    } else {
        $Broken.Add($Label)
        Show-Warn "$Label driver unavailable"
        Write-Log "Driver unavailable: $Label ($Module)"
        Write-Log $ImportResult.Output
    }
}

Show-Stage 95 "Creating dexcel command"

$DbToExcelPath = Join-Path $AppDir "db_to_excel.py"

try {
@"
@echo off
"$VenvPython" "$DbToExcelPath" %*
"@ | Set-Content -Encoding ASCII $LauncherPath

    Show-Ok "Launcher created at $LauncherPath"
} catch {
    Fail "Failed to create dexcel launcher. $_"
}

Show-Stage 100 "Finalizing installation"

$UserPath = [Environment]::GetEnvironmentVariable("Path", "User")

if ($UserPath -notlike "*$BinDir*") {
    if ([string]::IsNullOrWhiteSpace($UserPath)) {
        [Environment]::SetEnvironmentVariable("Path", $BinDir, "User")
    } else {
        [Environment]::SetEnvironmentVariable("Path", "$UserPath;$BinDir", "User")
    }

    $NeedNewTerminal = $true
    Show-Ok "Added Dexcel to user PATH"
} else {
    $NeedNewTerminal = $false
    Show-Ok "Dexcel command already exists in user PATH"
}

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host " Dexcel installed successfully" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Install path : $InstallDir"
Write-Host "Command      : dexcel"
Write-Host "Python       : $PythonVersion"
Write-Host "Log file     : $InstallLog"

if ($Working.Count -gt 0) {
    Write-Host ""
    Write-Host "Working database drivers:" -ForegroundColor Green

    foreach ($Item in $Working) {
        Write-Host "  [OK] $Item"
    }
}

if ($Broken.Count -gt 0) {
    Write-Host ""
    Write-Host "Unavailable database drivers:" -ForegroundColor Yellow

    foreach ($Item in $Broken) {
        Write-Host "  [--] $Item"
    }

    Write-Host ""
    Write-Host "These are optional. Dexcel can still run for the available drivers." -ForegroundColor Yellow
}

Write-Host ""

if ($NeedNewTerminal) {
    Write-Host "Open a NEW terminal, then run:" -ForegroundColor Yellow
    Write-Host "  dexcel"
} else {
    Write-Host "Run:" -ForegroundColor Yellow
    Write-Host "  dexcel"
}

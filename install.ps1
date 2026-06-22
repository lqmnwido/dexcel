# install.ps1 — Dexcel installer for Windows
#
# Install:
#   powershell -ExecutionPolicy Bypass -c "irm https://raw.githubusercontent.com/lqmnwido/dexcel/main/install.ps1 | iex"
#
# Run after install:
#   dexcel
# install.ps1 — Dexcel installer for Windows

$ErrorActionPreference = "Stop"

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

function Write-Log {
    param([string]$Message)
    Add-Content -Path $InstallLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

$script:CurrentStep = 0
$script:TotalSteps = 9
$script:ProgressWidth = 32

function Step {
    param([string]$Message)

    $script:CurrentStep++
    $Percent = [math]::Round(($script:CurrentStep / $script:TotalSteps) * 100)

    $Done = [math]::Floor(($Percent / 100) * $script:ProgressWidth)
    $Left = $script:ProgressWidth - $Done
    $Bar = ("#" * $Done) + ("-" * $Left)

    Write-Host ""
    Write-Host "[Dexcel Installer]" -ForegroundColor Cyan
    Write-Host "[$Bar] $Percent%"
    Write-Host "Step $($script:CurrentStep)/$($script:TotalSteps): $Message" -ForegroundColor White

    Write-Progress `
        -Activity "Dexcel Installer" `
        -Status "Step $($script:CurrentStep)/$($script:TotalSteps): $Message" `
        -PercentComplete $Percent

    Write-Log $Message
}

function Fail {
    param([string]$Message)

    Write-Progress -Activity "Dexcel Installer" -Completed
    Write-Host ""
    Write-Host "Installation failed." -ForegroundColor Red
    Write-Host $Message -ForegroundColor Red
    Write-Host "Log: $InstallLog" -ForegroundColor Yellow
    Write-Log "FATAL: $Message"
    exit 1
}

function Run-Command {
    param(
        [string]$Title,
        [scriptblock]$Command
    )

    Write-Log "RUN: $Title"

    try {
        & $Command 2>&1 | Tee-Object -FilePath $InstallLog -Append

        if ($LASTEXITCODE -ne 0) {
            throw "$Title failed with exit code $LASTEXITCODE"
        }
    } catch {
        Write-Log "ERROR: $_"
        throw $_
    }
}

function Test-PythonCommand {
    param([string]$Command)

    try {
        $Code = "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)"

        if ($Command -eq "py -3.12") {
            & py -3.12 -c $Code 2>$null
            return $LASTEXITCODE -eq 0
        }

        if ($Command -eq "py -3.11") {
            & py -3.11 -c $Code 2>$null
            return $LASTEXITCODE -eq 0
        }

        if ($Command -eq "py -3.10") {
            & py -3.10 -c $Code 2>$null
            return $LASTEXITCODE -eq 0
        }

        & $Command -c $Code 2>$null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
}

function Find-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "py -3.12") { return "py -3.12" }
        if (Test-PythonCommand "py -3.11") { return "py -3.11" }
        if (Test-PythonCommand "py -3.10") { return "py -3.10" }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "python") { return "python" }
    }

    return $null
}

function Create-Venv {
    param([string]$PythonCmd)

    if ($PythonCmd -eq "py -3.12") {
        Run-Command "Create venv using py -3.12" { py -3.12 -m venv $VenvDir }
    } elseif ($PythonCmd -eq "py -3.11") {
        Run-Command "Create venv using py -3.11" { py -3.11 -m venv $VenvDir }
    } elseif ($PythonCmd -eq "py -3.10") {
        Run-Command "Create venv using py -3.10" { py -3.10 -m venv $VenvDir }
    } else {
        Run-Command "Create venv using python" { python -m venv $VenvDir }
    }
}

Clear-Host

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host " Dexcel Installer" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Step "Checking compatible Python"

$PythonCmd = Find-Python

if (-not $PythonCmd) {
    Step "Installing Python 3.12"

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Run-Command "Install Python 3.12 using winget" {
            winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
        }
    } else {
        $PythonInstaller = Join-Path $env:TEMP "python-3.12-installer.exe"

        Invoke-WebRequest `
            -Uri "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe" `
            -OutFile $PythonInstaller

        Start-Process `
            -FilePath $PythonInstaller `
            -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_test=0" `
            -Wait

        Remove-Item $PythonInstaller -Force
    }

    $PythonCmd = Find-Python

    if (-not $PythonCmd) {
        Fail "Python 3.12 was installed but could not be found. Open a new terminal and re-run installer."
    }
}

Write-Host "Using Python: $PythonCmd" -ForegroundColor Green

Step "Checking old virtual environment"

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (Test-Path $VenvPython) {
    & $VenvPython -c "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)" 2>$null

    if ($LASTEXITCODE -ne 0) {
        Write-Host "Removing incompatible old virtual environment..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force $VenvDir
    }
}

Step "Downloading Dexcel files"

try {
    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/db_to_excel.py" `
        -OutFile (Join-Path $AppDir "db_to_excel.py")

    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/requirements.txt" `
        -OutFile (Join-Path $AppDir "requirements.txt")
} catch {
    Fail "Failed to download Dexcel files from $ReleaseBaseUrl. $_"
}

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

Step "Preparing Python environment"

if (-not (Test-Path $VenvPython)) {
    Create-Venv $PythonCmd

    if (-not (Test-Path $VenvPython)) {
        Fail "Failed to create virtual environment."
    }
} else {
    Write-Host "Existing compatible environment found." -ForegroundColor Green
}

Step "Upgrading installer tools"

try {
    Run-Command "Upgrade pip setuptools wheel" {
        & $VenvPython -m pip install --upgrade pip setuptools wheel
    }
} catch {
    Fail "Failed to upgrade pip."
}

Step "Installing Dexcel dependencies"

try {
    Run-Command "Install requirements.txt" {
        & $VenvPython -m pip install -r (Join-Path $AppDir "requirements.txt")
    }
} catch {
    Write-Host ""
    Write-Host "Manual debug command:" -ForegroundColor Yellow
    Write-Host "`"$VenvPython`" -m pip install -r `"$AppDir\requirements.txt`""

    Fail "Failed to install dependencies."
}

Step "Verifying core packages"

Run-Command "Verify core packages" {
    & $VenvPython -c "import pandas, openpyxl"
}

Step "Checking database drivers"

$Drivers = @{
    "MySQL/MariaDB" = "pymysql"
    "PostgreSQL" = "psycopg"
    "SQL Server" = "pyodbc"
    "Oracle" = "oracledb"
}

$Working = New-Object System.Collections.Generic.List[string]
$Broken = New-Object System.Collections.Generic.List[string]

foreach ($Label in $Drivers.Keys) {
    $Module = $Drivers[$Label]

    & $VenvPython -c "import $Module" *>> $InstallLog

    if ($LASTEXITCODE -eq 0) {
        $Working.Add($Label)
    } else {
        $Broken.Add($Label)
    }
}

Step "Creating dexcel command"

$DbToExcelPath = Join-Path $AppDir "db_to_excel.py"

@"
@echo off
"$VenvPython" "$DbToExcelPath" %*
"@ | Set-Content -Encoding ASCII $LauncherPath

$UserPath = [Environment]::GetEnvironmentVariable("Path", "User")

if ($UserPath -notlike "*$BinDir*") {
    [Environment]::SetEnvironmentVariable(
        "Path",
        "$UserPath;$BinDir",
        "User"
    )

    $NeedNewTerminal = $true
} else {
    $NeedNewTerminal = $false
}

Write-Progress -Activity "Dexcel Installer" -Completed

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host " Dexcel installed successfully" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Install path : $InstallDir"
Write-Host "Command      : dexcel"
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
}

Write-Host ""

if ($NeedNewTerminal) {
    Write-Host "Open a NEW terminal, then run:" -ForegroundColor Yellow
    Write-Host "  dexcel"
} else {
    Write-Host "Run:" -ForegroundColor Yellow
    Write-Host "  dexcel"
}

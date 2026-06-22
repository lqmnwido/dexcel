# install.ps1 — Dexcel installer for Windows
#
# Install:
#   powershell -ExecutionPolicy Bypass -c "irm https://raw.githubusercontent.com/lqmnwido/dexcel/main/install.ps1 | iex"
#
# Run after install:
#   dexcel


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

function Step {
    param(
        [int]$Percent,
        [string]$Message
    )

    Write-Progress -Activity "Installing Dexcel" -Status $Message -PercentComplete $Percent
    Write-Host "[$Percent%] $Message"
    Write-Log $Message
}

function Fail {
    param([string]$Message)

    Write-Progress -Activity "Installing Dexcel" -Completed

    Write-Host ""
    Write-Host "Error: $Message" -ForegroundColor Red
    Write-Host "Log: $InstallLog"
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
        & $Command *>> $InstallLog

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
        if ($Command -eq "py -3") {
            & py -3 -c "import sys; exit(0 if sys.version_info >= (3,8) else 1)" 2>$null
            return $LASTEXITCODE -eq 0
        }

        & $Command -c "import sys; exit(0 if sys.version_info >= (3,8) else 1)" 2>$null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
}

function Find-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "py -3") {
            return "py -3"
        }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "python") {
            return "python"
        }
    }

    return $null
}

Step 5 "Starting Dexcel installer..."

$PythonCmd = Find-Python

if (-not $PythonCmd) {
    Step 10 "Python 3.8+ not found. Installing Python..."

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Run-Command "Install Python using winget" {
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
        Fail "Python was installed but could not be found. Open a new terminal and re-run installer."
    }
}

Step 20 "Using Python: $PythonCmd"

Step 30 "Downloading Dexcel application files..."

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

if (-not (Test-Path $VenvPython)) {
    Step 40 "Creating isolated Python environment..."

    if ($PythonCmd -eq "py -3") {
        Run-Command "Create venv using py -3" {
            py -3 -m venv $VenvDir
        }
    } else {
        Run-Command "Create venv using python" {
            python -m venv $VenvDir
        }
    }

    if (-not (Test-Path $VenvPython)) {
        Fail "Failed to create virtual environment."
    }
} else {
    Step 40 "Existing Dexcel environment found. Updating it..."
}

Step 50 "Upgrading pip, setuptools, and wheel..."

try {
    Run-Command "Upgrade pip setuptools wheel" {
        & $VenvPython -m pip install --upgrade pip setuptools wheel
    }
} catch {
    Fail "Failed to upgrade pip. See log for details."
}

Step 60 "Installing Dexcel core dependencies..."

try {
    Run-Command "Install requirements.txt" {
        & $VenvPython -m pip install -r (Join-Path $AppDir "requirements.txt")
    }
} catch {
    Write-Host ""
    Write-Host "Dependency installation failed." -ForegroundColor Red
    Write-Host "Check this log:"
    Write-Host "  $InstallLog"
    Write-Host ""
    Write-Host "Manual debug command:"
    Write-Host "  `"$VenvPython`" -m pip install -r `"$AppDir\requirements.txt`""

    Fail "Failed to install dependencies."
}

Step 75 "Verifying core packages..."

Run-Command "Verify core packages" {
    & $VenvPython -c "import pandas, openpyxl"
}

Step 85 "Checking database drivers..."

$Drivers = @{
    "MySQL/MariaDB" = "pymysql"
    "PostgreSQL" = "psycopg2"
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

Step 92 "Creating dexcel command..."

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

Step 100 "Installation complete."
Write-Progress -Activity "Installing Dexcel" -Completed

Write-Host ""
Write-Host "Dexcel installed successfully." -ForegroundColor Green

if ($Working.Count -gt 0) {
    Write-Host "Working database drivers: $($Working -join ', ')"
}

if ($Broken.Count -gt 0) {
    Write-Host "Unavailable database drivers: $($Broken -join ', ')" -ForegroundColor Yellow
    Write-Host "See log: $InstallLog"
}

Write-Host ""

if ($NeedNewTerminal) {
    Write-Host "Open a NEW terminal, then run:"
    Write-Host "  dexcel"
} else {
    Write-Host "Run:"
    Write-Host "  dexcel"
}

# install.ps1 — Dexcel installer for Windows
#
# Install:
#   powershell -ExecutionPolicy Bypass -c "irm https://raw.githubusercontent.com/luqmanhafiz81/dexcel/main/install.ps1 | iex"
#
# Run after install:
#   dexcel

$ErrorActionPreference = "Stop"

$ReleaseBaseUrl = $env:DEXCEL_RELEASE_URL

if (-not $ReleaseBaseUrl) {
    $ReleaseBaseUrl = "https://github.com/luqmanhafiz81/dexcel/releases/latest/download"
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

    $Line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $InstallLog -Value $Line
}

function Info {
    param([string]$Message)

    Write-Host $Message
    Write-Log $Message
}

function Fail {
    param([string]$Message)

    Write-Host ""
    Write-Host "Error: $Message"
    Write-Host "Log: $InstallLog"
    Write-Log "FATAL: $Message"
    exit 1
}

function Test-PythonCommand {
    param([string]$Command)

    try {
        if ($Command -eq "py -3") {
            & py -3 -c "import sys; exit(0 if sys.version_info >= (3,8) else 1)" 2>$null
            return $LASTEXITCODE -eq 0
        } else {
            & $Command -c "import sys; exit(0 if sys.version_info >= (3,8) else 1)" 2>$null
            return $LASTEXITCODE -eq 0
        }
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

Info "Installing Dexcel..."

$PythonCmd = Find-Python

if (-not $PythonCmd) {
    Info "Python 3.8+ not found. Installing Python..."

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements | Tee-Object -FilePath $InstallLog -Append
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

Info "Using Python: $PythonCmd"

Info "Downloading Dexcel application files..."

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
    Info "Creating isolated Python environment..."

    if ($PythonCmd -eq "py -3") {
        & py -3 -m venv $VenvDir
    } else {
        & python -m venv $VenvDir
    }

    if (-not (Test-Path $VenvPython)) {
        Fail "Failed to create virtual environment."
    }
} else {
    Info "Existing Dexcel environment found. Updating it."
}

Info "Installing Dexcel dependencies..."

& $VenvPython -m pip install --upgrade pip *>> $InstallLog

if ($LASTEXITCODE -ne 0) {
    Fail "Failed to upgrade pip."
}

& $VenvPython -m pip install -r (Join-Path $AppDir "requirements.txt") *>> $InstallLog

if ($LASTEXITCODE -ne 0) {
    Fail "Failed to install dependencies."
}

Info "Verifying installation..."

& $VenvPython -c "import pandas, openpyxl" *>> $InstallLog

if ($LASTEXITCODE -ne 0) {
    Fail "Core packages failed verification."
}

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

Info "Creating dexcel command..."

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

Write-Host ""
Write-Host "Dexcel installed successfully."

if ($Working.Count -gt 0) {
    Write-Host "Working database drivers: $($Working -join ', ')"
}

if ($Broken.Count -gt 0) {
    Write-Host "Unavailable database drivers: $($Broken -join ', ')"
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
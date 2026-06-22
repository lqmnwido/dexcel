# install.ps1 - Dexcel installer for Windows
#
# Install:
#   powershell -ExecutionPolicy Bypass -c "irm https://dexcel.kohich.site/install.ps1 | iex"
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

$script:CurrentStep = 0
$script:TotalSteps = 10
$script:ProgressWidth = 34

function Write-Log {
    param([string]$Message)

    Add-Content -Path $InstallLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
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

function Show-Step {
    param([string]$Message)

    $script:CurrentStep++

    if ($script:CurrentStep -gt $script:TotalSteps) {
        $script:CurrentStep = $script:TotalSteps
    }

    $Percent = [math]::Round(($script:CurrentStep / $script:TotalSteps) * 100)
    $Done = [math]::Floor(($Percent / 100) * $script:ProgressWidth)
    $Left = $script:ProgressWidth - $Done

    $Bar = ("#" * $Done) + ("-" * $Left)

    Write-Host ""
    Write-Host "----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[$Bar] $Percent%" -ForegroundColor Cyan
    Write-Host "Step $($script:CurrentStep)/$($script:TotalSteps): $Message" -ForegroundColor White
    Write-Host "----------------------------------------------------" -ForegroundColor DarkGray

    Write-Log $Message
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

    return ($Arguments | ForEach-Object { Quote-Arg $_ }) -join " "
}

function Get-FileText {
    param([string]$Path)

    if (Test-Path $Path) {
        return Get-Content -Path $Path -Raw -ErrorAction SilentlyContinue
    }

    return ""
}

function Get-FileTail {
    param(
        [string]$Path,
        [int]$Lines = 40
    )

    if (Test-Path $Path) {
        return Get-Content -Path $Path -Tail $Lines -ErrorAction SilentlyContinue
    }

    return @()
}

function Invoke-Capture {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )

    $StdOut = Join-Path $env:TEMP ("dexcel_stdout_" + [guid]::NewGuid().ToString() + ".log")
    $StdErr = Join-Path $env:TEMP ("dexcel_stderr_" + [guid]::NewGuid().ToString() + ".log")

    try {
        $ArgumentLine = Join-Args $Arguments

        $Process = Start-Process `
            -FilePath $FilePath `
            -ArgumentList $ArgumentLine `
            -RedirectStandardOutput $StdOut `
            -RedirectStandardError $StdErr `
            -NoNewWindow `
            -PassThru

        $Process.WaitForExit()

        $Output = (Get-FileText $StdOut) + (Get-FileText $StdErr)

        return [pscustomobject]@{
            ExitCode = $Process.ExitCode
            Output   = $Output
        }
    } finally {
        Remove-Item $StdOut, $StdErr -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-LoggedCommand {
    param(
        [string]$Title,
        [string]$FilePath,
        [string[]]$Arguments
    )

    Write-Log "RUN: $Title"
    Write-Log "CMD: $FilePath $(Join-Args $Arguments)"

    $StdOut = Join-Path $env:TEMP ("dexcel_stdout_" + [guid]::NewGuid().ToString() + ".log")
    $StdErr = Join-Path $env:TEMP ("dexcel_stderr_" + [guid]::NewGuid().ToString() + ".log")

    try {
        $ArgumentLine = Join-Args $Arguments

        $Process = Start-Process `
            -FilePath $FilePath `
            -ArgumentList $ArgumentLine `
            -RedirectStandardOutput $StdOut `
            -RedirectStandardError $StdErr `
            -NoNewWindow `
            -PassThru

        $Spinner = @("|", "/", "-", "\")
        $SpinIndex = 0
        $StartTime = Get-Date

        while (-not $Process.HasExited) {
            $Spin = $Spinner[$SpinIndex % $Spinner.Count]
            $SpinIndex++

            $Elapsed = [int]((Get-Date) - $StartTime).TotalSeconds
            $Text = "  [$Spin] $Title... ${Elapsed}s"

            Write-Host -NoNewline "`r$Text"
            Start-Sleep -Milliseconds 180
        }

        $Process.WaitForExit()
        $ElapsedTotal = [int]((Get-Date) - $StartTime).TotalSeconds

        $OutText = Get-FileText $StdOut
        $ErrText = Get-FileText $StdErr

        if ($OutText) {
            Add-Content -Path $InstallLog -Value $OutText
        }

        if ($ErrText) {
            Add-Content -Path $InstallLog -Value $ErrText
        }

        $Clear = " " * 90
        Write-Host -NoNewline "`r$Clear`r"

        if ($Process.ExitCode -eq 0) {
            Write-Host "  [OK] $Title completed in ${ElapsedTotal}s" -ForegroundColor Green
            Write-Log "OK: $Title completed in ${ElapsedTotal}s"
            return
        }

        Write-Host "  [FAIL] $Title failed" -ForegroundColor Red

        Write-Host ""
        Write-Host "Last output:" -ForegroundColor Yellow

        $TailOut = Get-FileTail $StdOut 20
        $TailErr = Get-FileTail $StdErr 20

        foreach ($Line in $TailOut) {
            Write-Host "  $Line"
        }

        foreach ($Line in $TailErr) {
            Write-Host "  $Line"
        }

        throw "$Title failed with exit code $($Process.ExitCode)"
    } catch {
        Write-Log "ERROR: $_"
        throw $_
    } finally {
        Remove-Item $StdOut, $StdErr -Force -ErrorAction SilentlyContinue
    }
}

function Test-PythonCommand {
    param([string]$Command)

    $Code = "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)"

    try {
        if ($Command -eq "py -3.12") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.12", "-c", $Code)
            return $Result.ExitCode -eq 0
        }

        if ($Command -eq "py -3.11") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.11", "-c", $Code)
            return $Result.ExitCode -eq 0
        }

        if ($Command -eq "py -3.10") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.10", "-c", $Code)
            return $Result.ExitCode -eq 0
        }

        $Result = Invoke-Capture -FilePath $Command -Arguments @("-c", $Code)
        return $Result.ExitCode -eq 0
    } catch {
        return $false
    }
}

function Find-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "py -3.12") {
            return "py -3.12"
        }

        if (Test-PythonCommand "py -3.11") {
            return "py -3.11"
        }

        if (Test-PythonCommand "py -3.10") {
            return "py -3.10"
        }
    }

    if (Get-Command python -ErrorAction SilentlyContinue) {
        if (Test-PythonCommand "python") {
            return "python"
        }
    }

    return $null
}

function Get-PythonVersion {
    param([string]$PythonCmd)

    try {
        if ($PythonCmd -eq "py -3.12") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.12", "--version")
            return $Result.Output.Trim()
        }

        if ($PythonCmd -eq "py -3.11") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.11", "--version")
            return $Result.Output.Trim()
        }

        if ($PythonCmd -eq "py -3.10") {
            $Result = Invoke-Capture -FilePath "py" -Arguments @("-3.10", "--version")
            return $Result.Output.Trim()
        }

        $Result = Invoke-Capture -FilePath $PythonCmd -Arguments @("--version")
        return $Result.Output.Trim()
    } catch {
        return $PythonCmd
    }
}

function Create-Venv {
    param([string]$PythonCmd)

    if ($PythonCmd -eq "py -3.12") {
        Invoke-LoggedCommand "Create Python virtual environment" "py" @("-3.12", "-m", "venv", $VenvDir)
        return
    }

    if ($PythonCmd -eq "py -3.11") {
        Invoke-LoggedCommand "Create Python virtual environment" "py" @("-3.11", "-m", "venv", $VenvDir)
        return
    }

    if ($PythonCmd -eq "py -3.10") {
        Invoke-LoggedCommand "Create Python virtual environment" "py" @("-3.10", "-m", "venv", $VenvDir)
        return
    }

    Invoke-LoggedCommand "Create Python virtual environment" $PythonCmd @("-m", "venv", $VenvDir)
}

function Test-VenvCompatible {
    param([string]$PythonExe)

    if (-not (Test-Path $PythonExe)) {
        return $false
    }

    $Code = "import sys; exit(0 if sys.version_info >= (3,10) and sys.version_info < (3,13) else 1)"

    try {
        $Result = Invoke-Capture -FilePath $PythonExe -Arguments @("-c", $Code)
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
        $Result = Invoke-Capture -FilePath $PythonExe -Arguments @("-c", $Code)

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

Show-Step "Checking compatible Python"

$PythonCmd = Find-Python

if (-not $PythonCmd) {
    Show-Task "Compatible Python was not found. Installing Python 3.12..."

    try {
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Invoke-LoggedCommand `
                "Install Python 3.12 using winget" `
                "winget" `
                @(
                    "install",
                    "-e",
                    "--id",
                    "Python.Python.3.12",
                    "--accept-package-agreements",
                    "--accept-source-agreements"
                )
        } else {
            $PythonInstaller = Join-Path $env:TEMP "python-3.12-installer.exe"

            Show-Task "Downloading Python 3.12 installer..."

            Invoke-WebRequest `
                -Uri "https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe" `
                -OutFile $PythonInstaller

            Invoke-LoggedCommand `
                "Install Python 3.12" `
                $PythonInstaller `
                @("/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_test=0")

            Remove-Item $PythonInstaller -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Fail "Failed to install Python 3.12. $_"
    }

    $PythonCmd = Find-Python

    if (-not $PythonCmd) {
        Fail "Python 3.12 was installed but could not be found. Open a new terminal and re-run the installer."
    }
}

$PythonVersion = Get-PythonVersion $PythonCmd
Show-Ok "Using $PythonVersion via $PythonCmd"

Show-Step "Checking old virtual environment"

$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (Test-Path $VenvPython) {
    if (Test-VenvCompatible $VenvPython) {
        Show-Ok "Existing virtual environment is compatible"
    } else {
        Show-Warn "Old virtual environment is incompatible. Removing it..."
        Remove-Item -Recurse -Force $VenvDir
        Show-Ok "Old virtual environment removed"
    }
} else {
    Show-Task "No existing virtual environment found"
}

Show-Step "Downloading Dexcel files"

try {
    $MainFile = Join-Path $AppDir "db_to_excel.py"
    $RequirementsFile = Join-Path $AppDir "requirements.txt"

    Show-Task "Downloading db_to_excel.py"
    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/db_to_excel.py" `
        -OutFile $MainFile

    Show-Task "Downloading requirements.txt"
    Invoke-WebRequest `
        -Uri "$ReleaseBaseUrl/requirements.txt" `
        -OutFile $RequirementsFile

    Show-Ok "Application files downloaded"
} catch {
    Fail "Failed to download Dexcel files from $ReleaseBaseUrl. $_"
}

Show-Step "Preparing Python environment"

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

Show-Step "Upgrading installer tools"

try {
    Invoke-LoggedCommand `
        "Upgrade pip, setuptools, and wheel" `
        $VenvPython `
        @("-m", "pip", "install", "--upgrade", "pip", "setuptools", "wheel")
} catch {
    Fail "Failed to upgrade pip, setuptools, and wheel. $_"
}

Show-Step "Installing Dexcel dependencies"

try {
    Invoke-LoggedCommand `
        "Install Python dependencies" `
        $VenvPython `
        @("-m", "pip", "install", "-r", $RequirementsFile)
} catch {
    Write-Host ""
    Write-Host "Manual debug command:" -ForegroundColor Yellow
    Write-Host "  `"$VenvPython`" -m pip install -r `"$RequirementsFile`""

    Fail "Failed to install Dexcel dependencies. $_"
}

Show-Step "Verifying core packages"

try {
    Invoke-LoggedCommand `
        "Verify pandas and openpyxl" `
        $VenvPython `
        @("-c", "import pandas, openpyxl")
} catch {
    Fail "Core packages failed verification. $_"
}

Show-Step "Checking database drivers"

$Drivers = @{
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

Show-Step "Creating dexcel command"

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

Show-Step "Finalizing installation"

$UserPath = [Environment]::GetEnvironmentVariable("Path", "User")

if ($UserPath -notlike "*$BinDir*") {
    [Environment]::SetEnvironmentVariable(
        "Path",
        "$UserPath;$BinDir",
        "User"
    )

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

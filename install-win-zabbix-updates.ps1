<#
  All-in-one installer for Windows Updates monitoring with Zabbix Agent 2.
  - Ensures PSWindowsUpdate module is installed
  - Ensures Git for Windows is installed (winget -> choco -> standalone)
  - Deploys UserParameters + PowerShell script
  - Restarts "Zabbix Agent 2"
  - Emits SUCCESS/ERROR log lines with timestamps

  Edit the two RAW URLs below if you keep files in a different repo or path.
#>

Param(
  [string]$Ps1Url = "https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/windows-updates.ps1",
  [string]$ConfUrl = "https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/windows-updates.conf"
)

$ErrorActionPreference = "Stop"

function TS { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
function OK($msg) { Write-Host ("SUCCESS | {0} | {1}" -f (TS), $msg) -ForegroundColor Green }
function ERR($msg){ Write-Host ("ERROR   | {0} | {1}" -f (TS), $msg) -ForegroundColor Red }

#--- Admin check -----------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { ERR "Please run this PowerShell as Administrator."; exit 1 }

try {
  #--- Ensure folders ------------------------------------------------------
  $ZbxBase = "C:\Program Files\Zabbix Agent 2"
  $ZbxConfD = Join-Path $ZbxBase "zabbix_agent2.d"
  $ZbxScripts = Join-Path $ZbxBase "scripts"

  if (-not (Test-Path $ZbxBase))   { ERR "Zabbix Agent 2 base folder not found: $ZbxBase"; exit 1 }
  if (-not (Test-Path $ZbxConfD))  { New-Item -ItemType Directory -Path $ZbxConfD -Force | Out-Null }
  if (-not (Test-Path $ZbxScripts)){ New-Item -ItemType Directory -Path $ZbxScripts -Force | Out-Null }
  OK "Zabbix Agent 2 folders verified."

  #--- Ensure PSGallery & NuGet -------------------------------------------
  $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
  if (-not $repo) {
    Register-PSRepository -Default
    OK "Registered PSGallery."
  } else {
    OK "PSGallery already registered."
  }

  if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    OK "NuGet package provider installed."
  } else {
    OK "NuGet package provider present."
  }

  #--- Install PSWindowsUpdate --------------------------------------------
  if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    try {
      Install-Module -Name PSWindowsUpdate -Force -AllowClobber | Out-Null
      OK "Installed PSWindowsUpdate (system scope)."
    } catch {
      Install-Module -Name PSWindowsUpdate -Scope CurrentUser -Force -AllowClobber | Out-Null
      OK "Installed PSWindowsUpdate (CurrentUser scope)."
    }
  } else {
    OK "PSWindowsUpdate already installed."
  }

  # Import & sanity check
  Import-Module PSWindowsUpdate -Force
  if (Get-Command Get-WindowsUpdate -ErrorAction SilentlyContinue) {
    OK "Get-WindowsUpdate available."
  } else {
    ERR "PSWindowsUpdate imported but Get-WindowsUpdate not found."
  }

  #--- Ensure Git for Windows ---------------------------------------------
  function Test-Git { return [bool](Get-Command git.exe -ErrorAction SilentlyContinue) }

  if (-not (Test-Git)) {
    $installed = $false

    # Try winget
    if (Get-Command winget.exe -ErrorAction SilentlyContinue) {
      try {
        winget install --id Git.Git -e --source winget --silent --accept-source-agreements --accept-package-agreements
        Start-Sleep -Seconds 3
        if (Test-Git) { $installed = $true; OK "Git installed via winget." }
      } catch { }
    }

    # Try chocolatey
    if (-not $installed -and (Get-Command choco.exe -ErrorAction SilentlyContinue)) {
      try {
        choco install git -y --no-progress
        Start-Sleep -Seconds 3
        if (Test-Git) { $installed = $true; OK "Git installed via chocolatey." }
      } catch { }
    }

    # Fallback: standalone installer
    if (-not $installed) {
      try {
        $tmp = New-Item -ItemType Directory -Path (Join-Path $env:TEMP ("git-inst-" + [guid]::NewGuid().ToString())) -Force
        $gitExe = Join-Path $tmp.FullName "Git-64-bit.exe"
        Invoke-WebRequest -Uri "https://github.com/git-for-windows/git/releases/latest/download/Git-64-bit.exe" -OutFile $gitExe
        & $gitExe /VERYSILENT /NORESTART | Out-Null
        Start-Sleep -Seconds 5
        if (Test-Git) { $installed = $true; OK "Git installed via standalone installer." }
        Remove-Item -Path $tmp.FullName -Recurse -Force -ErrorAction SilentlyContinue
      } catch {
        ERR "Failed to install Git via fallback: $($_.Exception.Message)"
      }
    }

    if (-not $installed) {
      ERR "Could not install Git automatically. Please install Git for Windows and re-run."
    }
  } else {
    OK "Git is present."
  }

  #--- Download monitoring files ------------------------------------------
  $ps1Dst  = Join-Path $ZbxScripts "windows-updates.ps1"
  $confDst = Join-Path $ZbxConfD   "windows-updates.conf"

  Invoke-WebRequest -Uri $Ps1Url  -OutFile $ps1Dst
  Invoke-WebRequest -Uri $ConfUrl -OutFile $confDst
  OK "Downloaded monitoring files to Zabbix Agent 2 directories."

  #--- Restart Zabbix Agent 2 ---------------------------------------------
  $svc = Get-Service -Name "Zabbix Agent 2" -ErrorAction Stop
  if ($svc.Status -eq 'Running') {
    Restart-Service -Name "Zabbix Agent 2" -Force
  } else {
    Start-Service -Name "Zabbix Agent 2"
  }
  # Wait a moment and verify
  Start-Sleep -Seconds 2
  $svc.Refresh()
  if ($svc.Status -eq 'Running') {
    OK "Zabbix Agent 2 restarted successfully."
  } else {
    ERR "Zabbix Agent 2 is not running after restart."
    exit 1
  }

  OK "All tasks completed."
}
catch {
  ERR $_.Exception.Message
  exit 1
}

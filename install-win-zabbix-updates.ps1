<#
  All-in-one installer for Windows Updates monitoring with Zabbix Agent 2.

  FINAL FIX:
  - NEVER imports PSWindowsUpdate (GPO-safe)
  - Uses COM-only monitoring (already proven working)
  - Chocolatey forced before use
  - Git fallback via GitHub API (always latest)
  - Terminal stays open on errors
#>

Param(
  [string]$Ps1Url  = "https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/windows-updates.ps1",
  [string]$ConfUrl = "https://raw.githubusercontent.com/pthoelken/windowsupdates-zabbix-monitoring/refs/heads/main/windows-updates.conf"
)

$ErrorActionPreference = "Stop"

function TS { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
function OK($m)   { Write-Host ("SUCCESS | {0} | {1}" -f (TS), $m) -ForegroundColor Green }
function WARN($m) { Write-Host ("WARN    | {0} | {1}" -f (TS), $m) -ForegroundColor Yellow }
function FAIL($m) {
  Write-Host ("ERROR   | {0} | {1}" -f (TS), $m) -ForegroundColor Red
  Write-Host ""
  Read-Host "Press ENTER to exit"
  exit 1
}

# --- Admin check ----------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal] `
  [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
  FAIL "Please run this PowerShell session as Administrator."
}

try {
  # --- Zabbix folders -----------------------------------------------------
  $ZbxBase     = "C:\Program Files\Zabbix Agent 2"
  $ZbxConfD   = Join-Path $ZbxBase "zabbix_agent2.d"
  $ZbxScripts = Join-Path $ZbxBase "scripts"

  if (-not (Test-Path $ZbxBase)) {
    FAIL "Zabbix Agent 2 not found at $ZbxBase"
  }

  New-Item -ItemType Directory -Path $ZbxConfD   -Force | Out-Null
  New-Item -ItemType Directory -Path $ZbxScripts -Force | Out-Null
  OK "Zabbix Agent 2 directories verified."

  # --- OPTIONAL: PSWindowsUpdate install (NO IMPORT) ----------------------
  if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    try {
      Install-Module PSWindowsUpdate -Force -AllowClobber | Out-Null
      OK "PSWindowsUpdate installed (not imported, GPO-safe)."
    } catch {
      WARN "PSWindowsUpdate could not be installed (non-fatal)."
    }
  } else {
    OK "PSWindowsUpdate already present (not imported)."
  }

  # --- Git detection ------------------------------------------------------
  function Test-Git { [bool](Get-Command git.exe -ErrorAction SilentlyContinue) }

  if (-not (Test-Git)) {
    $installed = $false

    # winget
    if (Get-Command winget.exe -ErrorAction SilentlyContinue) {
      try {
        winget install --id Git.Git -e --silent `
          --accept-source-agreements --accept-package-agreements
        Start-Sleep 3
        if (Test-Git) { OK "Git installed via winget."; $installed = $true }
      } catch { WARN "winget Git install failed." }
    }

    # Chocolatey bootstrap
    if (-not $installed) {
      if (-not (Get-Command choco.exe -ErrorAction SilentlyContinue)) {
        OK "Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-Expression ((New-Object Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Start-Sleep 5
      }

      if (Get-Command choco.exe -ErrorAction SilentlyContinue) {
        try {
          choco install git -y --no-progress
          Start-Sleep 3
          if (Test-Git) { OK "Git installed via Chocolatey."; $installed = $true }
        } catch { WARN "Chocolatey Git install failed." }
      }
    }

    # GitHub API fallback
    if (-not $installed) {
      OK "Using GitHub API fallback."
      $tmp = New-Item -ItemType Directory -Path (Join-Path $env:TEMP ("git-inst-" + [guid]::NewGuid())) -Force
      try {
        $rel = Invoke-RestMethod "https://api.github.com/repos/git-for-windows/git/releases/latest" `
          -Headers @{ "User-Agent"="zabbix-installer" }
        $asset = $rel.assets | Where-Object { $_.name -match '^Git-.*-64-bit\.exe$' } | Select-Object -First 1
        if (-not $asset) { FAIL "No Git 64-bit installer found." }

        $exe = Join-Path $tmp.FullName $asset.name
        Invoke-WebRequest $asset.browser_download_url -OutFile $exe
        & $exe /VERYSILENT /NORESTART | Out-Null
        Start-Sleep 5

        if (-not (Test-Git)) { FAIL "Git installer ran but git.exe not found." }
        OK ("Git installed ({0})." -f $asset.name)
      }
      finally {
        Remove-Item $tmp.FullName -Recurse -Force -ErrorAction SilentlyContinue
      }
    }
  } else {
    OK "Git already present."
  }

  # --- Deploy monitoring files -------------------------------------------
  Invoke-WebRequest $Ps1Url  -OutFile (Join-Path $ZbxScripts "windows-updates.ps1")
  Invoke-WebRequest $ConfUrl -OutFile (Join-Path $ZbxConfD "windows-updates.conf")
  OK "Monitoring files deployed."

  # --- Restart Zabbix Agent 2 --------------------------------------------
  Restart-Service "Zabbix Agent 2" -Force
  Start-Sleep 2

  if ((Get-Service "Zabbix Agent 2").Status -ne 'Running') {
    FAIL "Zabbix Agent 2 did not start."
  }

  OK "Zabbix Agent 2 restarted successfully."
  OK "Installation completed successfully."
  Read-Host "Press ENTER to exit"
}
catch {
  FAIL $_.Exception.Message
}
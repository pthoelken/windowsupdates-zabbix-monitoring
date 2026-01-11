<#
  All-in-one installer for Windows Updates monitoring with Zabbix Agent 2.
  FIXED:
  - Ensures Chocolatey is installed before using it
  - Keeps PowerShell window open on errors
  - Git fallback always uses latest release via GitHub API
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

  # --- PSGallery / NuGet --------------------------------------------------
  if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
    Register-PSRepository -Default
    OK "PSGallery registered."
  }

  if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    OK "NuGet provider installed."
  }

  # --- PSWindowsUpdate ----------------------------------------------------
  if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-Module PSWindowsUpdate -Force -AllowClobber | Out-Null
    OK "PSWindowsUpdate installed."
  } else {
    OK "PSWindowsUpdate already present."
  }

  Import-Module PSWindowsUpdate -Force
  if (-not (Get-Command Get-WindowsUpdate -ErrorAction SilentlyContinue)) {
    FAIL "PSWindowsUpdate installed but Get-WindowsUpdate not available."
  }
  OK "Get-WindowsUpdate verified."

  # --- Git detection ------------------------------------------------------
  function Test-Git { [bool](Get-Command git.exe -ErrorAction SilentlyContinue) }

  if (-not (Test-Git)) {
    $installed = $false

    # --- winget -----------------------------------------------------------
    if (Get-Command winget.exe -ErrorAction SilentlyContinue) {
      try {
        winget install --id Git.Git -e --silent `
          --accept-source-agreements --accept-package-agreements
        Start-Sleep 3
        if (Test-Git) {
          OK "Git installed via winget."
          $installed = $true
        }
      } catch {
        WARN "winget Git install failed."
      }
    }

    # --- Chocolatey bootstrap (FORCE) ------------------------------------
    if (-not $installed) {
      if (-not (Get-Command choco.exe -ErrorAction SilentlyContinue)) {
        OK "Chocolatey not found. Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-Expression ((New-Object Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Start-Sleep 5
      }

      if (Get-Command choco.exe -ErrorAction SilentlyContinue) {
        try {
          choco install git -y --no-progress
          Start-Sleep 3
          if (Test-Git) {
            OK "Git installed via Chocolatey."
            $installed = $true
          }
        } catch {
          WARN "Chocolatey Git install failed."
        }
      } else {
        WARN "Chocolatey installation failed."
      }
    }

    # --- Fallback: GitHub API (ALWAYS LATEST) ------------------------------
    if (-not $installed) {
      OK "Using standalone GitHub installer fallback."

      $tmp = New-Item -ItemType Directory `
        -Path (Join-Path $env:TEMP ("git-inst-" + [guid]::NewGuid())) -Force

      try {
        $release = Invoke-RestMethod `
          -Uri "https://api.github.com/repos/git-for-windows/git/releases/latest" `
          -Headers @{ "User-Agent" = "zabbix-installer" }

        $asset = $release.assets |
          Where-Object { $_.name -match '^Git-.*-64-bit\.exe$' } |
          Select-Object -First 1

        if (-not $asset) {
          FAIL "No 64-bit Git installer found in latest GitHub release."
        }

        $gitExe = Join-Path $tmp.FullName $asset.name
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $gitExe
        & $gitExe /VERYSILENT /NORESTART | Out-Null
        Start-Sleep 5

        if (Test-Git) {
          OK ("Git installed via standalone installer ({0})." -f $asset.name)
          $installed = $true
        } else {
          FAIL "Git installer executed but git.exe not found afterwards."
        }
      }
      finally {
        Remove-Item -Path $tmp.FullName -Recurse -Force -ErrorAction SilentlyContinue
      }
    }

    if (-not $installed) {
      FAIL "Git could not be installed by any method."
    }
  } else {
    OK "Git already present."
  }

  # --- Deploy monitoring files -------------------------------------------
  $ps1Dst  = Join-Path $ZbxScripts "windows-updates.ps1"
  $confDst = Join-Path $ZbxConfD   "windows-updates.conf"

  Invoke-WebRequest -Uri $Ps1Url  -OutFile $ps1Dst
  Invoke-WebRequest -Uri $ConfUrl -OutFile $confDst
  OK "Monitoring files deployed."

  # --- Restart Zabbix Agent 2 --------------------------------------------
  $svc = Get-Service "Zabbix Agent 2" -ErrorAction Stop
  if ($svc.Status -eq 'Running') {
    Restart-Service "Zabbix Agent 2" -Force
  } else {
    Start-Service "Zabbix Agent 2"
  }

  Start-Sleep 2
  $svc.Refresh()
  if ($svc.Status -ne 'Running') {
    FAIL "Zabbix Agent 2 failed to start."
  }

  OK "Zabbix Agent 2 restarted successfully."
  OK "Installation completed successfully."
  Write-Host ""
  Read-Host "Press ENTER to exit"
}
catch {
  FAIL $_.Exception.Message
}
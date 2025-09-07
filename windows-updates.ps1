Param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Security','NonSecurity','All','Reboot')]
    [string]$Mode
)

# Return integers only, no extra output
$ErrorActionPreference = 'Stop'
function Write-Result($value) {
    [Console]::Out.Write($value)
    exit 0
}
function Write-ErrorAndExit($code) {
    [Console]::Out.Write($code)
    exit 0
}

# Reboot check does not need PSWindowsUpdate
if ($Mode -eq 'Reboot') {
    $rebootKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    if (Test-Path $rebootKey) { Write-Result 1 } else { Write-Result 0 }
}

# Try to import PSWindowsUpdate; if missing, return -1
try {
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-ErrorAndExit -1
    }
    Import-Module PSWindowsUpdate -ErrorAction Stop | Out-Null
} catch {
    Write-ErrorAndExit -1
}

try {
    # Query available (not yet installed) updates
    $updates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput -AcceptAll -IgnoreReboot | Where-Object { -not $_.IsInstalled }
    $allCount = ($updates | Measure-Object).Count

    $securityCount = ($updates | Where-Object {
        ($_ | Select-Object -ExpandProperty Categories -ErrorAction SilentlyContinue) -match 'Security' -or
        ($_.Title -match 'Security')
    } | Measure-Object).Count

    switch ($Mode) {
        'All'        { Write-Result $allCount }
        'Security'   { Write-Result $securityCount }
        'NonSecurity'{ Write-Result ([math]::Max(0, $allCount - $securityCount)) }
    }
} catch {
    Write-ErrorAndExit -2
}

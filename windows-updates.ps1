Param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Security','NonSecurity','All','Reboot')]
    [string]$Mode
)

# Silence everything except our final integer
$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$InformationPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'
$PSModuleAutoLoadingPreference = 'None'

function Out-Int([int]$v){ [Console]::Out.Write($v); exit 0 }

# --- Fast path: Reboot flag (kein Cmdlet, nur .NET) ---
if ($Mode -eq 'Reboot') {
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Default
        )
        $sub  = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
        if ($null -ne $sub) { Out-Int 1 } else { Out-Int 0 }
    } catch { Out-Int 0 }
}

# --- Logging nur bei Fehlern (keine Ausgabe an Zabbix) ---
$DiagFile = 'C:\ProgramData\Zabbix\winupdates_diag.log'
function LogDiag($msg) {
    try {
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $line = "$ts | $msg"
        New-Item -ItemType Directory -Path ([System.IO.Path]::GetDirectoryName($DiagFile)) -Force -ErrorAction SilentlyContinue | Out-Null
        Add-Content -Path $DiagFile -Value $line -Encoding UTF8
    } catch { }
}

# --- Dienste sicherstellen (Starttyp Disabled -> Manual), ohne Cmdlets
function Ensure-Service($name){
    try {
        & sc.exe config $name start= demand *>$null
        & sc.exe start  $name *>$null
    } catch {
        LogDiag "Service ${name}: $($_.Exception.Message)"
    }
}
Ensure-Service 'wuauserv'
Ensure-Service 'BITS'

# --- COM-Suche (robust, mehrere Versuche) ---
function Get-CountsViaCOM {
    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()

        # Policy (per Registry, ohne Cmdlets)
        $lm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Default)
        $auKey = $lm.OpenSubKey('SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU')
        $wuKey = $lm.OpenSubKey('SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate')
        $useWUServer = 0; $wuServer = $null
        if ($auKey) { $useWUServer = ($auKey.GetValue('UseWUServer',0)) }
        if ($wuKey) { $wuServer    =  $wuKey.GetValue('WUServer',$null) }

        if ($useWUServer -eq 1 -and [string]::IsNullOrWhiteSpace($wuServer)) {
            LogDiag "Policy enforces WSUS (UseWUServer=1) aber WUServer ist leer."
        }

        $serverSelections = if ($useWUServer -eq 1) { @(1,2,0) } else { @(2,1,0) }  # 1=WSUS, 2=WindowsUpdate, 0=Default
        $criteriaList     = @(
            "IsInstalled=0 and IsHidden=0 and Type='Software'",
            "IsInstalled=0 and IsHidden=0"
        )
        $onlineFlags      = @($true, $false)

        foreach($online in $onlineFlags){
            try { $searcher.Online = $online } catch { LogDiag "Set Online=${online}: $($_.Exception.Message)" }
            foreach($sel in $serverSelections){
                try { $searcher.ServerSelection = $sel } catch { LogDiag "ServerSelection=${sel}: $($_.Exception.Message)" }
                foreach($crit in $criteriaList){
                    try {
                        $res = $searcher.Search($crit)
                        if ($res -and $res.Updates -ne $null) {
                            $cnt = [int]$res.Updates.Count
                            $all = $cnt; $sec = 0
                            for ($i=0; $i -lt $cnt; $i++) {
                                $u = $res.Updates.Item($i)
                                $isSec = $false
                                if ($u.Categories) {
                                    foreach ($c in $u.Categories) {
                                        if ($c -and $c.Name -match 'Security') { $isSec = $true; break }
                                    }
                                }
                                if (-not $isSec -and $u.Title -match 'Security') { $isSec = $true }
                                if ($isSec) { $sec++ }
                            }
                            return @{ All = [int]$all; Sec = [int]$sec }
                        }
                    } catch {
                        LogDiag "Search fail (Online=${online}, Sel=${sel}, Crit='$crit'): $($_.Exception.Message)"
                    }
                }
            }
        }
        return $null
    } catch {
        LogDiag "COM init/searcher error: $($_.Exception.Message)"
        return $null
    }
}

$counts = Get-CountsViaCOM

# --- Finale Ausgabe: nur Zahlen; bei Fehler -> 0 ---
if ($null -eq $counts) {
    LogDiag "Alle Versuche fehlgeschlagen. Rückgabe 0 für Zabbix."
    switch ($Mode) {
        'All'        { Out-Int 0 }
        'Security'   { Out-Int 0 }
        'NonSecurity'{ Out-Int 0 }
    }
}

switch ($Mode) {
    'All'        { Out-Int $counts.All }
    'Security'   { Out-Int $counts.Sec }
    'NonSecurity'{ Out-Int ([math]::Max(0, $counts.All - $counts.Sec)) }
}

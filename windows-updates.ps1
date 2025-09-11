Param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Security','NonSecurity','All','Reboot')]
    [string]$Mode
)

function Out-Int([int]$v){ [Console]::Out.Write($v); exit 0 }

# Keine Autoloads/Cmdlets voraussetzen
$PSModuleAutoLoadingPreference = 'None'
$ErrorActionPreference = 'Stop'

# .NET-Logging (kein Add-Content)
$Diag = 'C:\Program Files\Zabbix Agent 2\winupdates_diag.log'
function LogDiag([string]$m){
    try{
        $ts = [DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss')
        $dir = [System.IO.Path]::GetDirectoryName($Diag)
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        [System.IO.File]::AppendAllText($Diag, ("{0} | {1}`r`n" -f $ts, $m), [System.Text.Encoding]::UTF8)
    }catch{}
}

# Reboot-Flag (nur .NET)
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

# Dienste sicherstellen (nur sc.exe)
function Ensure-Service([string]$name){
    try {
        & sc.exe config $name start= demand *> $null
        & sc.exe start  $name          *> $null
    } catch {
        LogDiag( ("Service {0}: {1}" -f $name, $_.Exception.Message) )
    }
}
Ensure-Service 'wuauserv'
Ensure-Service 'BITS'

# Policy aus Registry (nur .NET)
function Get-Policy(){
    $r = @{ UseWUServer = 0; WUServer = $null }
    try {
        $lm  = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Default)
        $au  = $lm.OpenSubKey('SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU')
        $wu  = $lm.OpenSubKey('SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate')
        if ($au) { $r.UseWUServer = [int]($au.GetValue('UseWUServer',0)) }
        if ($wu) { $r.WUServer    = [string]($wu.GetValue('WUServer',$null)) }
    } catch { }
    return $r
}

# COM-Suche mit Retries (ohne Cmdlets)
function Get-CountsViaCOM {
    try {
        # USO-Scan anstoßen
        try { & usoclient.exe StartScan *> $null } catch {}

        # COM-Objekte mit .NET Activator (kein New-Object)
        $typeSession = [type]::GetTypeFromProgID('Microsoft.Update.Session')
        if ($null -eq $typeSession) { LogDiag('ProgID Microsoft.Update.Session nicht verfügbar.'); return @{ All = 0; Sec = 0 } }
        $session  = [System.Activator]::CreateInstance($typeSession)
        # ClientApplicationID setzen (hilft in einigen Umgebungen)
        try { $session.ClientApplicationID = 'Zabbix-WindowsUpdates' } catch {}

        $searcher = $session.CreateUpdateSearcher()

        # Optionen (falls vorhanden)
        try { $searcher.Online = $true } catch {}
        try { $searcher.IncludePotentiallySupersededUpdates = $true } catch {}

        $pol = Get-Policy
        if ($pol.UseWUServer -eq 1 -and [string]::IsNullOrWhiteSpace($pol.WUServer)) {
            LogDiag('UseWUServer=1, aber WUServer leer.')
        }

        # 1=WSUS, 2=WindowsUpdate, 0=Default
        $serverSelections = if ($pol.UseWUServer -eq 1) { @(1,2,0) } else { @(2,1,0) }
        # Breitere Kriterien
        $criteriaList = @(
            "IsInstalled=0 and IsHidden=0 and Type='Software'",
            "IsInstalled=0 and IsHidden=0",
            "IsInstalled=0"
        )
        $onlineFlags = @($true, $false)

        # 18 Versuche x 5s = 90s
        for ($attempt = 1; $attempt -le 18; $attempt++) {
            foreach ($online in $onlineFlags) {
                try { $searcher.Online = $online } catch { LogDiag( ("Online={0}: {1}" -f $online, $_.Exception.Message) ) }
                foreach ($sel in $serverSelections) {
                    try { $searcher.ServerSelection = $sel } catch { LogDiag( ("ServerSelection={0}: {1}" -f $sel, $_.Exception.Message) ) }
                    foreach ($crit in $criteriaList) {
                        try {
                            $res = $searcher.Search($crit)
                            if ($res -and $res.Updates -ne $null) {
                                $cnt = [int]$res.Updates.Count
                                if ($cnt -gt 0) {
                                    $all = $cnt; $sec = 0
                                    for ($i=0; $i -lt $cnt; $i++) {
                                        $u = $res.Updates.Item($i)

                                        $isSec = $false
                                        # Kategorien prüfen
                                        try {
                                            $cats = $u.Categories
                                            if ($cats) {
                                                for ($j=0; $j -lt $cats.Count; $j++) {
                                                    $c = $cats.Item($j)
                                                    if ($c -and $c.Name -match 'Security') { $isSec = $true; break }
                                                }
                                            }
                                        } catch {}

                                        # Fallback: Titel enthält "Security"
                                        if (-not $isSec) {
                                            try { if ($u.Title -match 'Security') { $isSec = $true } } catch {}
                                        }

                                        if ($isSec) { $sec++ }
                                    }
                                    return @{ All = [int]$all; Sec = [int]$sec }
                                }
                            }
                        } catch {
                            LogDiag( ("Search fail (Try={0}, Online={1}, Sel={2}, Crit='{3}'): {4}" -f $attempt, $online, $sel, $crit, $_.Exception.Message) )
                        }
                    }
                }
            }
            [System.Threading.Thread]::Sleep(5000)
        }
        # Nach allen Versuchen: 0 zurück
        return @{ All = 0; Sec = 0 }
    } catch {
        LogDiag( "COM error: " + $_.Exception.Message )
        return @{ All = 0; Sec = 0 }
    }
}

$counts = Get-CountsViaCOM

switch ($Mode) {
    'All'        { Out-Int $counts.All }
    'Security'   { Out-Int $counts.Sec }
    'NonSecurity'{ Out-Int ([Math]::Max(0, $counts.All - $counts.Sec)) }
}

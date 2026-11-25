<#
================================================================================
 Remove-MassLynx  -  uninstall Waters MassLynx V4.2 from this computer
================================================================================

 A small tool for lab PCs. Double-click "Remove-MassLynx.cmd" (or start this
 script); a window appears with the current state and three buttons:

    Controleren            alleen kijken, er wordt niets veranderd
    MassLynx verwijderen   MassLynx en de resten weghalen
    Logbestand openen      het verslag van de laatste keer bekijken

 Without a window (for IT or scripting):

    powershell -ExecutionPolicy Bypass -File .\Remove-MassLynx.ps1 -Action Check
    powershell -ExecutionPolicy Bypass -File .\Remove-MassLynx.ps1 -Action Remove

 What it removes: MassLynx V4.2, MassLynx Data, Pesticide Data, Dioxin Data,
 TargetLynx and the spectrometer driver - plus the leftovers that belong to them
 (program folder, COM registrations, services, shortcuts, registry values).

 What it leaves alone: all other Waters software, including shared components
 and instrument drivers. They are recorded before the removal and
 checked afterwards; if one of them disappeared the tool reports it as an error.

 Nothing is installed. Everything is logged in
 C:\ProgramData\MassLynxOnderhoud\Logs\ (one file per run, plus the full
 Windows Installer log per product).
================================================================================
#>

[CmdletBinding()]
param(
    [ValidateSet('Gui','Check','Remove')]
    [string] $Action = 'Gui',

    [string] $LogDir = 'C:\ProgramData\MassLynxOnderhoud\Logs',

    [switch] $NoBackup,
    [switch] $Quiet
)

$ErrorActionPreference = 'Continue'
$ProgressPreference    = 'SilentlyContinue'

# ---------------------------------------------------------------- vaste waarden
$script:Version     = '2.1.1'
$script:InstallRoot = 'C:\MassLynx'
$script:SettingsDir = 'C:\ProgramData\MassLynxOnderhoud'

# Where MassLynx itself lives. Only paths inside these folders are cleaned up:
# a file, service or registry value that points elsewhere does not belong to us.
$script:OwnedPaths = @(
    'C:\MassLynx\',
    'C:\TargetLynx\',
    'C:\Program Files (x86)\Waters Instruments\',
    'C:\Program Files\Waters Instruments\'
)

# Productcodes van MassLynx V4.2 (inclusief de onderdelen en de spectrometer-driver).
$script:MassLynxCodes = @(
    '{4595B43D-1C31-4809-93C7-7136F41E92D7}',   # Waters MassLynx V4.2 SCN1051
    '{57A3B98E-EF23-461C-9C76-8898DEFECEF6}',   # Waters MassLynx V4.2 Data
    '{9988C6A4-68EE-4896-AFEA-80AD55E8AA4A}',   # Waters Pesticide V4.2 Data
    '{CD861CD5-0A1B-48D8-80C9-68BE04CCA5E2}',   # Waters Dioxin V4.2 Data
    '{9AED080D-F466-45FE-A200-BDC617FC70AF}',   # Waters TargetLynx
    '{7122AF7F-31D8-45F0-A155-2963DD094768}'    # Waters Xevo TQ-XS Driver Module
)

# Name matching for other versions/service releases of the same family.
$script:MassLynxNamePattern = 'MassLynx V4\.2|Waters TargetLynx|Pesticide V4\.2 Data|Dioxin V4\.2 Data'

# Program identifiers (ProgIDs) that MassLynx registers itself.
$script:MassLynxProgIdPattern = '^(MassLynx|TargetLynx|MLynx)\.'

# ---------------------------------------------------------------- toestand
$script:OthersBefore = $null
$script:DoBackup     = $true

# ---------------------------------------------------------------- gui-variabelen
$script:Form        = $null
$script:LogBox      = $null
$script:StateBox    = $null
$script:StatusLabel = $null
$script:Progress    = $null
$script:ChkBackup   = $null
$script:Buttons     = @()
$script:LogFile     = $null

# ================================================================ hulpfuncties
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string] $Text,
        [ValidateSet('INFO','STAP','OK','LET OP','FOUT')][string] $Level = 'INFO'
    )

    $stamp = Get-Date -Format 'HH:mm:ss'
    $short = $Text
    if ($Level -eq 'STAP') { $short = '- ' + $Text }
    if ($Level -eq 'OK')   { $short = '  ok  ' + $Text }
    if ($Level -eq 'LET OP') { $short = '  !!  ' + $Text }
    if ($Level -eq 'FOUT') { $short = '  XX  ' + $Text }

    if ($script:LogFile) {
        $full = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $Level.PadRight(6) + '  ' + $Text
        Add-Content -Path $script:LogFile -Value $full -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    if ($script:LogBox) {
        $script:LogBox.AppendText($stamp + '  ' + $short + "`r`n")
        $script:LogBox.SelectionStart = $script:LogBox.TextLength
        $script:LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    elseif (-not $Quiet) {
        $color = 'Gray'
        # Only colours that really exist in System.ConsoleColor.
        if ($Level -eq 'OK')     { $color = 'Green' }
        if ($Level -eq 'LET OP') { $color = 'Yellow' }
        if ($Level -eq 'FOUT')   { $color = 'Red' }
        if ($Level -eq 'STAP')   { $color = 'Cyan' }
        try { Write-Host ($stamp + '  ' + $short) -ForegroundColor $color }
        catch { Write-Host ($stamp + '  ' + $short) }
    }

    if ($script:StatusLabel -and $Level -eq 'STAP') {
        $script:StatusLabel.Text = $Text
        $script:StatusLabel.Refresh()
    }
}

function Write-Header {
    param([string]$Text)
    Write-Log ''
    Write-Log ('=== ' + $Text + ' ===') 'STAP'
}

function Write-ErrorDetail {
    <#  Schrijft niet alleen de melding maar ook de regel en de stacktrace in het log,
        zodat een fout door IT (of door de maker) te herleiden is. #>
    param($ErrorRecord)
    try {
        $msg  = $ErrorRecord.Exception.Message
        $line = $ErrorRecord.InvocationInfo.ScriptLineNumber
        $code = ''
        if ($ErrorRecord.InvocationInfo.Line) { $code = $ErrorRecord.InvocationInfo.Line.Trim() }
        Write-Log ('STOP: ' + $msg) 'FOUT'
        if ($line) { Write-Log ('   regel ' + $line + ': ' + $code) 'FOUT' }
        if ($ErrorRecord.ScriptStackTrace) {
            foreach ($frame in ($ErrorRecord.ScriptStackTrace -split "`n")) {
                if ($frame.Trim()) { Write-Log ('   ' + $frame.Trim()) 'FOUT' }
            }
        }
        if ($ErrorRecord.Exception.InnerException) {
            Write-Log ('   oorzaak: ' + $ErrorRecord.Exception.InnerException.Message) 'FOUT'
        }
    } catch {
        Write-Log 'STOP: onbekende fout' 'FOUT'
    }
}

function Test-Admin {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Hide-OwnConsole {
    <#  Hide the PowerShell window so the operator only sees the application window. #>
    param([switch]$Always)
    if (-not $Always) {
        try {
            if ((Get-Process -Id $PID).MainWindowTitle -ne 'MassLynx verwijderen') { return }
        } catch { }
    }
    try {
        if (-not ('ConsoleVenster' -as [type])) {
            Add-Type -Namespace W -Name ConsoleVenster -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        $h = [W.ConsoleVenster]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][W.ConsoleVenster]::ShowWindow($h, 0) }
    } catch { }
}

function Get-RegValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
    if ($item) { return $item.$Name }
    return $null
}

function Test-OwnedServerPath {
    <#  Hoort dit pad (van een COM-server, dienst of koppeling) bij MassLynx?  #>
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $v = $Value.Trim()
    if ($v.StartsWith('"')) {
        $rest = $v.Substring(1)
        $i = $rest.IndexOf('"')
        if ($i -gt 0) { $v = $rest.Substring(0, $i) } else { $v = $rest }
    } else {
        $m = [regex]::Match($v, '^(.*?\.(exe|dll))(\s|,|$)')
        if ($m.Success) { $v = $m.Groups[1].Value }
    }
    foreach ($p in $script:OwnedPaths) {
        if ($v.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------- producten
function Test-OwnedProduct {
    param($Product)
    if (-not $Product) { return $false }
    $code = '{' + ([string]$Product.PSChildName).Trim('{','}') + '}'
    if ($script:MassLynxCodes -contains $code) { return $true }
    if ($Product.DisplayName -match $script:MassLynxNamePattern) { return $true }
    if ($Product.InstallSource -and ($Product.InstallSource -match 'MassLynx4\.2_SCN1051|MassLynx V4\.2')) { return $true }
    return $false
}

function Get-InstalledProducts {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    return @(Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName } |
        Select-Object DisplayName, DisplayVersion, InstallSource, PSChildName, Publisher)
}

function Get-MassLynxProducts {
    return @(Get-InstalledProducts | Where-Object { Test-OwnedProduct $_ } | Sort-Object DisplayName -Unique)
}

function Get-OtherWatersProducts {
    <#  Waters software that does NOT belong to MassLynx.  #>
    return @(Get-InstalledProducts |
        Where-Object { (($_.Publisher -match 'Waters|Micromass') -or
                        ($_.DisplayName -match 'Waters|MassLynx|TargetLynx|ACQUITY')) -and
                       (-not (Test-OwnedProduct $_)) } |
        Sort-Object DisplayName -Unique)
}

function Get-MassLynxProcesses {
    return @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ExecutablePath -and (Test-OwnedServerPath $_.ExecutablePath) })
}

function Get-MassLynxServices {
    return @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
        Where-Object { $_.PathName -and (Test-OwnedServerPath $_.PathName) })
}

function Get-InstallState {
    $products = Get-MassLynxProducts
    $others   = Get-OtherWatersProducts
    $micromass = 'HKLM:\SOFTWARE\WOW6432Node\Micromass'

    return [pscustomobject]@{
        Installed      = ($products.Count -gt 0)
        Products       = $products
        Others         = $others
        Security       = (Get-RegValue -Path $micromass -Name 'Security')
        SecurityMode   = (Get-RegValue -Path $micromass -Name 'SecurityMode')
        InstrumentName = (Get-RegValue -Path $micromass -Name 'InstrumentName')
        InstallRoot    = (Test-Path $script:InstallRoot)
        Running        = ((Get-MassLynxProcesses).Count -gt 0)
    }
}

function Format-StateSummary {
    param($State)

    $lines = @()
    if ($State.Installed) {
        $lines += "MassLynx       : AANWEZIG ($($State.Products.Count) onderdelen)"
        $names = ($State.Products | ForEach-Object { ($_.DisplayName -replace '^Waters ', '') + ' ' + $_.DisplayVersion }) -join ', '
        $lines += '                 ' + $names
    } else {
        $lines += 'MassLynx       : NIET AANWEZIG'
    }

    if ($State.Installed) {
        $sec = 'aan'
        if (($State.Security -eq $false) -or ($State.Security -eq 0) -or ($State.Security -eq 'FALSE')) { $sec = 'uit' }
        $lines += 'Beveiliging    : ' + $sec + '   (Security=' + $State.Security + ', mode=' + $State.SecurityMode + ')'
        if ($State.InstrumentName) { $lines += 'Spectrometer   : INGESTELD - ' + $State.InstrumentName }
        else { $lines += 'Spectrometer   : geen ingeschakeld' }
    }

    if ($State.Running) { $lines += 'MassLynx nu    : geopend' } else { $lines += 'MassLynx nu    : niet geopend' }
    $lines += 'Programmamap   : ' + $script:InstallRoot + ' ' + $(if ($State.InstallRoot) { 'aanwezig' } else { 'niet aanwezig' })

    if ($State.Others.Count -gt 0) {
        $lines += "Andere Waters  : $($State.Others.Count) onderdeel/onderdelen blijven staan"
    } else {
        $lines += 'Andere Waters  : geen andere Waters-software gevonden'
    }
    return ($lines -join "`r`n")
}

function Invoke-MsiExec {
    param(
        [Parameter(Mandatory = $true)][string[]] $Arguments,
        [Parameter(Mandatory = $true)][string]   $LogPath,
        [string] $What = 'msiexec',
        [int]    $TimeoutMinutes = 25
    )

    $exe = Join-Path $env:WINDIR 'System32\msiexec.exe'
    $proc = Start-Process -FilePath $exe -ArgumentList $Arguments -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $ticks = 0
    $lastShown = ''

    while (-not $proc.HasExited) {
        Start-Sleep -Milliseconds 800
        if ($script:LogBox) { [System.Windows.Forms.Application]::DoEvents() }
        $ticks++

        if ((Get-Date) -gt $deadline) {
            Write-Log ($What + ' reageert niet meer na ' + $TimeoutMinutes + ' minuten - afgebroken') 'FOUT'
            try { $proc.Kill() } catch { }
            Start-Sleep -Seconds 3
            return -1
        }

        if ($ticks % 12 -eq 0) {
            $tail = ''
            try { $tail = (Get-Content -Path $LogPath -Tail 1 -Encoding Unicode -ErrorAction SilentlyContinue) } catch { }
            if ($tail -and $tail -ne $lastShown) {
                if ($tail.Length -gt 120) { $tail = $tail.Substring(0, 120) }
                Write-Log ('   bezig: ' + $tail.Trim())
                $lastShown = $tail
            } else {
                Write-Log '   bezig...'
            }
        }
    }

    return $proc.ExitCode
}

# ================================================================ verwijderen
function Stop-MassLynxEverything {
    Write-Log "MassLynx en hulp-programma's afsluiten" 'STAP'

    # Only MassLynx programs. Other Waters software is left alone.
    $names = @('mlynx4','MassLynx','TargetLynx','TargetMethodEditor','LancasterUI',
               'TargetRegistrySysmon','SecurityEditor','CapLC','MLynxSle')
    foreach ($n in $names) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Log ('   afsluiten: ' + $_.ProcessName + ' (pid ' + $_.Id + ')')
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }

    Get-MassLynxProcesses | ForEach-Object {
        Write-Log ('   afsluiten: ' + $_.Name + ' (pid ' + $_.ProcessId + ') ' + $_.ExecutablePath)
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 3

    Get-MassLynxServices | ForEach-Object {
        Write-Log ('   dienst stoppen: ' + $_.Name + ' (' + $_.State + ')')
        Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    Get-Process msiexec, ISBEW64 -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Log ('   achtergebleven ' + $_.ProcessName + ' gestopt')
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
}

function Backup-MassLynxData {
    if ($NoBackup -or ($script:DoBackup -eq $false)) {
        Write-Log 'Veiligheidskopie overgeslagen (niet aangevinkt)' 'LET OP'
        return $null
    }
    if (-not (Test-Path $script:InstallRoot)) { return $null }

    Write-Log 'Projecten en data veiligstellen' 'STAP'
    $size = (Get-ChildItem $script:InstallRoot -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    if (-not $size) { $size = 0 }
    $free = (Get-PSDrive ($script:InstallRoot.Substring(0,1))).Free
    Write-Log ('   grootte ' + [math]::Round($size / 1MB) + ' MB, vrij op schijf ' + [math]::Round($free / 1GB, 1) + ' GB')

    if ($free -lt ($size * 2 + 500MB)) {
        Write-Log 'Onvoldoende schijfruimte voor een veiligheidskopie - doorgegaan zonder back-up' 'LET OP'
        return $null
    }

    $dest = 'C:\MassLynx-backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    $rc = (Start-Process -FilePath 'robocopy.exe' -Wait -PassThru -WindowStyle Hidden -ArgumentList @(
        ('"' + $script:InstallRoot + '"'), ('"' + $dest + '"'),
        '/E', '/R:1', '/W:1', '/NFL', '/NDL', '/NJH', '/NJS', '/NP')).ExitCode

    if ($rc -ge 8) {
        Write-Log ('Veiligheidskopie mislukt (robocopy ' + $rc + ') - doorgegaan zonder back-up') 'LET OP'
        return $null
    }

    $count = (Get-ChildItem $dest -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    Write-Log ('   veiligheidskopie gemaakt: ' + $dest + ' (' + $count + ' bestanden)') 'OK'
    return $dest
}

function Invoke-MlRemoval {
    param([string]$WorkLogDir)

    $products = Get-MassLynxProducts
    if ($products.Count -eq 0) {
        Write-Log 'Geen MassLynx-onderdelen gevonden om te verwijderen' 'OK'
        return @()
    }

    $others = Get-OtherWatersProducts
    if ($others.Count -gt 0) {
        Write-Log ('Deze Waters-software blijft staan (hoort niet bij MassLynx): ' +
                   (($others | ForEach-Object { $_.DisplayName }) -join ', ')) 'LET OP'
    }

    Write-Header ("MassLynx-onderdelen verwijderen ($($products.Count))")
    foreach ($p in $products) { Write-Log ('   gevonden: ' + $p.DisplayName + '  ' + $p.DisplayVersion) }

    # Hoofdprogramma als laatste, driver en datapakketten eerst.
    $main    = @($products | Where-Object { $_.DisplayName -match 'MassLynx V4\.2 SCN' })
    $rest    = @($products | Where-Object { $main -notcontains $_ } |
                Sort-Object @{ Expression = { if ($_.DisplayName -match 'Driver Module') { 0 } else { 1 } } }, DisplayName)
    $ordered = @($rest) + @($main)

    $results = @()
    foreach ($p in $ordered) {
        if (-not $p) { continue }
        $id  = $p.PSChildName.Trim('{','}')
        $log = Join-Path $WorkLogDir ('verwijderen-' + $id + '.log')
        Write-Log ('verwijderen: ' + $p.DisplayName) 'STAP'

        $code = Invoke-MsiExec -Arguments @('/x', $p.PSChildName, '/qn', '/norestart', 'REBOOT=ReallySuppress', '/l*v', ('"' + $log + '"')) `
                               -LogPath $log -What $p.DisplayName

        if (($code -ne 0) -and ($code -ne 3010) -and ($code -ne 1605)) {
            Write-Log ('   eerste poging gaf ' + $code + ' - nog een keer proberen') 'LET OP'
            Get-Process msiexec -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $log2 = Join-Path $WorkLogDir ('verwijderen-' + $id + '-poging2.log')
            $code = Invoke-MsiExec -Arguments @('/x', $p.PSChildName, '/qn', '/norestart', 'REBOOT=ReallySuppress', '/l*v', ('"' + $log2 + '"')) `
                                   -LogPath $log2 -What $p.DisplayName
        }

        if (($code -eq 0) -or ($code -eq 3010) -or ($code -eq 1605)) {
            Write-Log ('   klaar (resultaat ' + $code + ')') 'OK'
        } else {
            Write-Log ('   mislukt (resultaat ' + $code + ') - zie ' + $log) 'FOUT'
        }

        $results += [pscustomobject]@{ Product = $p.DisplayName; Code = $code }
        Start-Sleep -Seconds 2
    }

    $left = Get-MassLynxProducts
    if ($left.Count -eq 0) { Write-Log 'Alle MassLynx-onderdelen zijn verwijderd' 'OK' }
    else { Write-Log ('Nog geregistreerd: ' + (($left | ForEach-Object { $_.DisplayName }) -join ', ')) 'LET OP' }

    return $results
}

function Invoke-MlCleanup {
    param([string]$WorkLogDir)

    Write-Header 'Resten van MassLynx opruimen'

    foreach ($n in @('LancasterUI','TargetRegistrySysmon','mlynx4')) {
        Get-Process -Name $n -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Log ('   afsluiten: ' + $_.ProcessName)
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 3

    foreach ($d in @($script:InstallRoot, 'C:\TargetLynx')) {
        if (-not (Test-Path $d)) { continue }
        $m = Get-ChildItem $d -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum
        Write-Log ('   map verwijderen: ' + $d + ' (' + $m.Count + ' bestanden, ' + [math]::Round($m.Sum / 1MB) + ' MB)')
        Remove-Item $d -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $d) { cmd /c ('rd /s /q "' + $d + '"') 2>&1 | Out-Null }
        if (Test-Path $d) { Write-Log ('   kon niet volledig weg: ' + $d) 'LET OP' }
    }

    # Waters Instruments is shared: only remove it when nothing of another Waters
    # application is left in it.
    foreach ($d in @('C:\Program Files (x86)\Waters Instruments', 'C:\Program Files\Waters Instruments')) {
        if (-not (Test-Path $d)) { continue }
        $rest = @(Get-ChildItem $d -Recurse -ErrorAction SilentlyContinue)
        if ($rest.Count -eq 0) {
            Remove-Item $d -Force -ErrorAction SilentlyContinue
            Write-Log ('   lege map verwijderd: ' + $d) 'OK'
        } else {
            Write-Log ('   blijft staan (ook door andere Waters-software gebruikt): ' + $d) 'OK'
        }
    }

    Write-Log '   register en COM-koppelingen van MassLynx opschonen'
    $views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    $removedClsid = 0
    $removedProgId = 0

    foreach ($view in $views) {
        try {
            $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        } catch {
            Write-Log ('   registerweergave ' + $view + ' niet te openen') 'LET OP'
            continue
        }

        $clsids = $hklm.OpenSubKey('Software\Classes\CLSID', $true)
        if ($clsids) {
            foreach ($name in $clsids.GetSubKeyNames()) {
                foreach ($sub in @('LocalServer32','InprocServer32')) {
                    $k = $clsids.OpenSubKey(('' + $name + '\' + $sub))
                    if ($k) {
                        $value = [string]$k.GetValue('')
                        $k.Close()
                        if (Test-OwnedServerPath $value) {
                            try { $clsids.DeleteSubKeyTree($name); $removedClsid++ } catch { }
                            break
                        }
                    }
                }
            }
            $clsids.Close()
        }

        $classes = $hklm.OpenSubKey('Software\Classes', $true)
        if ($classes) {
            $prefixed = @($classes.GetSubKeyNames() | Where-Object { $_ -match $script:MassLynxProgIdPattern })
            foreach ($p in $prefixed) {
                try { $classes.DeleteSubKeyTree($p); $removedProgId++ } catch { }
            }
            $classes.Close()
        }
        $hklm.Close()
    }
    Write-Log ('   ' + $removedClsid + ' COM-koppelingen en ' + $removedProgId + ' programma-koppelingen van MassLynx verwijderd') 'OK'

    # Registry: only the MassLynx branch and the values MassLynx writes itself.
    $micromass = 'HKLM:\SOFTWARE\WOW6432Node\Micromass'
    if (Test-Path $micromass) {
        $mlKey = Join-Path $micromass 'MassLynx'
        if (Test-Path $mlKey) {
            Remove-Item $mlKey -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log '   Micromass\MassLynx verwijderd' 'OK'
        }
        foreach ($v in @('Security','SecurityMode','SecurityOn','InstrumentName','InstrumentExe',
                         'InstrumentDirectory','InstrumentPath','InstrumentType')) {
            if ($null -ne (Get-RegValue -Path $micromass -Name $v)) {
                Remove-ItemProperty -Path $micromass -Name $v -Force -ErrorAction SilentlyContinue
            }
        }
        $rest = @(Get-ChildItem $micromass -ErrorAction SilentlyContinue)
        if ($rest.Count -eq 0) {
            Remove-Item $micromass -Force -ErrorAction SilentlyContinue
            Write-Log '   lege map Micromass verwijderd' 'OK'
        }
    }

    # Instrument settings left behind by MassLynx. The rest of the Waters key stays
    # (other Waters software uses it too).
    foreach ($k in @('HKLM:\SOFTWARE\WOW6432Node\Waters\Instruments\XEVO-TQXS',
                     'HKLM:\SOFTWARE\WOW6432Node\Waters\Instruments\LANCASTER',
                     'HKLM:\SOFTWARE\Waters\Instruments\XEVO-TQXS',
                     'HKLM:\SOFTWARE\Waters\Instruments\LANCASTER')) {
        if (Test-Path $k) {
            Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log ('   instrumentgegevens verwijderd: ' + $k) 'OK'
        }
    }

    $gac = 'C:\Windows\assembly\GAC_MSIL\MassLynxSecurity'
    if (Test-Path $gac) {
        Remove-Item $gac -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log '   beveiligingsbibliotheek uit de GAC verwijderd' 'OK'
    }

    Get-MassLynxServices | ForEach-Object {
        Write-Log ('   dienst verwijderen: ' + $_.Name)
        sc.exe delete $_.Name 2>&1 | Out-Null
    }

    Get-ChildItem "$env:ProgramData\Microsoft\Windows\Start Menu" -Recurse -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match 'MassLynx|TargetLynx' } | ForEach-Object {
            Write-Log ('   snelkoppeling verwijderen: ' + $_.FullName)
            Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        }
}

function Test-MlClean {
    Write-Header 'Controleren of alles van MassLynx weg is'

    $results = @()
    $add = {
        param($Name, $Ok, $Detail)
        $script:MlCleanResults += [pscustomobject]@{ Name = $Name; Ok = [bool]$Ok; Detail = $Detail }
    }
    $script:MlCleanResults = @()

    $products = Get-MassLynxProducts
    & $add 'geen MassLynx-onderdelen meer geregistreerd' ($products.Count -eq 0) (($products | ForEach-Object { $_.DisplayName }) -join ', ')

    $msiLeft = @()
    $ud = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products'
    if (Test-Path $ud) {
        foreach ($key in Get-ChildItem $ud -ErrorAction SilentlyContinue) {
            $ip = Join-Path $key.PSPath 'InstallProperties'
            if (Test-Path $ip) {
                $dn = Get-RegValue -Path $ip -Name 'DisplayName'
                $pc = ([string]$key.PSChildName).Trim()
                $isOwned = ($dn -and ($dn -match $script:MassLynxNamePattern)) -or
                           ($script:MassLynxCodes -contains ('{' + $pc + '}'))
                if ($isOwned) { $msiLeft += $dn }
            }
        }
    }
    & $add 'geen resten in de Windows Installer-database' ($msiLeft.Count -eq 0) ($msiLeft -join ', ')

    & $add ('programmamap weg: ' + $script:InstallRoot) (-not (Test-Path $script:InstallRoot)) ''
    & $add 'registergegevens Micromass\MassLynx weg' (-not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Micromass\MassLynx')) ''
    & $add 'geen spectrometer meer ingesteld' (-not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Waters\Instruments\XEVO-TQXS') -and
                                               -not (Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Waters\Instruments\LANCASTER')) ''

    $leftClsid = 0
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        try { $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view) } catch { continue }
        $clsids = $hklm.OpenSubKey('Software\Classes\CLSID')
        if ($clsids) {
            foreach ($name in $clsids.GetSubKeyNames()) {
                foreach ($sub in @('LocalServer32','InprocServer32')) {
                    $k = $clsids.OpenSubKey(('' + $name + '\' + $sub))
                    if ($k) {
                        $value = [string]$k.GetValue('')
                        $k.Close()
                        if (Test-OwnedServerPath $value) { $leftClsid++ }
                    }
                }
            }
            $clsids.Close()
        }
        $hklm.Close()
    }
    & $add 'geen programma-koppelingen naar MassLynx meer' ($leftClsid -eq 0) ($leftClsid.ToString() + ' gevonden')
    & $add 'beveiligingsbibliotheek niet meer aanwezig' (-not (Test-Path 'C:\Windows\assembly\GAC_MSIL\MassLynxSecurity')) ''
    & $add 'geen MassLynx-diensten meer' ((Get-MassLynxServices).Count -eq 0) ''
    & $add "geen MassLynx-programma's meer actief" ((Get-MassLynxProcesses).Count -eq 0) ''

    # Whatever other Waters software was installed before must still be there.
    if ($null -ne $script:OthersBefore -and $script:OthersBefore.Count -gt 0) {
        $now = @(Get-OtherWatersProducts | ForEach-Object { $_.DisplayName })
        $missing = @($script:OthersBefore | Where-Object { $now -notcontains $_ })
        & $add 'andere Waters-software niet aangetast' ($missing.Count -eq 0) ('verdwenen: ' + ($missing -join ', '))
    }

    foreach ($r in $script:MlCleanResults) {
        if ($r.Ok) { Write-Log ('   ok    ' + $r.Name) 'OK' }
        else       { Write-Log ('   NIET  ' + $r.Name + '  ' + $r.Detail) 'FOUT' }
    }

    $fails = @($script:MlCleanResults | Where-Object { -not $_.Ok })
    if ($fails.Count -eq 0) { Write-Log ('Alles van MassLynx is weg - ' + $script:MlCleanResults.Count + ' controles geslaagd') 'OK' }
    else { Write-Log ("$($fails.Count) van $($script:MlCleanResults.Count) controles niet in orde") 'FOUT' }

    return $script:MlCleanResults
}

# ================================================================ acties
function Invoke-ActionCheck {
    $state = Get-InstallState
    Write-Log ''
    Write-Log 'Status van deze computer:' 'STAP'
    (Format-StateSummary -State $state) -split "`r?`n" | ForEach-Object { Write-Log ('   ' + $_) }

    if ($state.Installed) {
        Write-Log ''
        $svc = Get-MassLynxServices | Where-Object { $_.State -eq 'Running' }
        if ($svc.Count -gt 0) { Write-Log ('   diensten draaien: ' + (($svc | ForEach-Object { $_.Name }) -join ', ')) 'OK' }
        if (Test-Path (Join-Path $script:InstallRoot 'mlynx4.exe')) { Write-Log '   MassLynx-programma aanwezig' 'OK' }
        else { Write-Log '   mlynx4.exe ontbreekt' 'LET OP' }
        if (Test-Path (Join-Path $script:InstallRoot 'TargetLynx.exe')) { Write-Log '   TargetLynx aanwezig' 'OK' }
    } else {
        Test-MlClean | Out-Null
    }
    Write-Log ''
    Write-Log ('Logbestand: ' + $script:LogFile)
}

function Invoke-ActionRemove {
    Write-Header 'MassLynx verwijderen'
    $script:OthersBefore = @(Get-OtherWatersProducts | ForEach-Object { $_.DisplayName })
    $workLogDir = Split-Path -Parent $script:LogFile

    Stop-MassLynxEverything
    Backup-MassLynxData | Out-Null
    Invoke-MlRemoval -WorkLogDir $workLogDir | Out-Null
    Invoke-MlCleanup -WorkLogDir $workLogDir
    return (Test-MlClean)
}

# ================================================================ venster
function Initialize-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'MassLynx verwijderen ' + $script:Version
    $form.ClientSize = New-Object System.Drawing.Size(820, 560)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $script:Form = $form

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'MassLynx verwijderen'
    $title.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
    $title.SetBounds(16, 10, 500, 30)
    $form.Controls.Add($title)

    $sub = New-Object System.Windows.Forms.Label
    $sub.Text = 'Haalt MassLynx V4.2 van deze computer. Andere Waters-software blijft staan.'
    $sub.SetBounds(16, 42, 788, 20)
    $sub.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($sub)

    $groupState = New-Object System.Windows.Forms.GroupBox
    $groupState.Text = 'Status van deze computer'
    $groupState.SetBounds(16, 66, 788, 130)
    $form.Controls.Add($groupState)

    $script:StateBox = New-Object System.Windows.Forms.TextBox
    $script:StateBox.Multiline = $true
    $script:StateBox.ReadOnly = $true
    $script:StateBox.ScrollBars = 'Vertical'
    $script:StateBox.Font = New-Object System.Drawing.Font('Consolas', 8)
    $script:StateBox.SetBounds(12, 20, 764, 100)
    $script:StateBox.BackColor = [System.Drawing.Color]::White
    $script:StateBox.TabStop = $false
    $groupState.Controls.Add($script:StateBox)

    $script:ChkBackup = New-Object System.Windows.Forms.CheckBox
    $script:ChkBackup.Text = 'Projecten en data in C:\MassLynx eerst veiligstellen (aanbevolen)'
    $script:ChkBackup.SetBounds(16, 204, 500, 22)
    $script:ChkBackup.Checked = $true
    $form.Controls.Add($script:ChkBackup)

    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = 'Controleren'
    $btnCheck.SetBounds(16, 234, 150, 34)
    $form.Controls.Add($btnCheck)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'MassLynx verwijderen'
    $btnRemove.SetBounds(176, 234, 250, 34)
    $btnRemove.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnRemove)

    $btnLog = New-Object System.Windows.Forms.Button
    $btnLog.Text = 'Logbestand openen'
    $btnLog.SetBounds(436, 234, 160, 34)
    $form.Controls.Add($btnLog)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Sluiten'
    $btnClose.SetBounds(684, 234, 120, 34)
    $form.Controls.Add($btnClose)

    $script:Buttons = @($btnCheck, $btnRemove, $btnLog, $btnClose, $script:ChkBackup)

    $script:Progress = New-Object System.Windows.Forms.ProgressBar
    $script:Progress.SetBounds(16, 278, 788, 12)
    $script:Progress.Style = 'Continuous'
    $form.Controls.Add($script:Progress)

    $script:LogBox = New-Object System.Windows.Forms.TextBox
    $script:LogBox.Multiline = $true
    $script:LogBox.ReadOnly = $true
    $script:LogBox.ScrollBars = 'Vertical'
    $script:LogBox.Font = New-Object System.Drawing.Font('Consolas', 8)
    $script:LogBox.SetBounds(16, 298, 788, 222)
    $script:LogBox.BackColor = [System.Drawing.Color]::White
    $script:LogBox.TabStop = $false
    $form.Controls.Add($script:LogBox)

    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Text = 'Gereed.'
    $script:StatusLabel.SetBounds(16, 530, 788, 20)
    $form.Controls.Add($script:StatusLabel)

    $script:RefreshState = {
        $state = Get-InstallState
        $script:StateBox.Text = Format-StateSummary -State $state
        $script:StateBox.Select(0, 0)
    }

    $script:RunBlock = {
        param([string]$Mode)

        $script:DoBackup = [bool]$script:ChkBackup.Checked
        foreach ($b in $script:Buttons) { $b.Enabled = $false }
        $script:Progress.Style = 'Marquee'
        $script:Progress.MarqueeAnimationSpeed = 30
        $script:StatusLabel.Text = 'Bezig...'
        $script:LogBox.Clear()
        $ok = $true
        $message = ''

        try {
            switch ($Mode) {
                'Check'  { Invoke-ActionCheck }
                'Remove' { $r = Invoke-ActionRemove; if (@($r | Where-Object { -not $_.Ok }).Count -gt 0) { $ok = $false } }
            }
        } catch {
            $ok = $false
            $message = $_.Exception.Message
            Write-ErrorDetail -ErrorRecord $_
        }

        $script:Progress.Style = 'Continuous'
        $script:Progress.Value = 0
        foreach ($b in $script:Buttons) { $b.Enabled = $true }
        & $script:RefreshState

        if ($Mode -eq 'Check') {
            $script:StatusLabel.Text = 'Controle klaar.'
            return
        }

        if ($ok) {
            $script:StatusLabel.Text = 'Klaar - MassLynx is weg.'
            $script:Progress.Value = 100
            [System.Windows.Forms.MessageBox]::Show(
                ("Klaar.`r`n`r`nMassLynx is van deze computer verwijderd en de resten zijn opgeruimd.`r`n`r`nHet volledige verslag staat in:`r`n" + $script:LogFile),
                'MassLynx verwijderen', 'OK', 'Information') | Out-Null
        } else {
            $script:StatusLabel.Text = 'Niet gelukt - zie de melding.'
            [System.Windows.Forms.MessageBox]::Show(
                ("Het is niet gelukt om alles af te ronden.`r`n`r`n" + $message + "`r`n`r`nNeem contact op met IT en geef dit logbestand mee:`r`n" + $script:LogFile),
                'MassLynx verwijderen', 'OK', 'Error') | Out-Null
        }
    }

    & $script:RefreshState

    $btnCheck.Add_Click({ & $script:RunBlock 'Check' })

    $btnRemove.Add_Click({
        $state = Get-InstallState
        $extra = ''
        if ($state.Running) { $extra = "`r`n`r`nLet op: MassLynx is nu geopend. Het wordt afgesloten; werk dat nog niet is opgeslagen gaat verloren." }
        $backupText = 'Er wordt eerst een veiligheidskopie van C:\MassLynx gemaakt.'
        if (-not $script:ChkBackup.Checked) { $backupText = 'Er wordt GEEN veiligheidskopie gemaakt.' }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            ("MassLynx wordt van deze computer verwijderd.`r`n`r`n" + $backupText +
             "`r`nAndere Waters-software blijft staan.`r`n`r`nDaarna is MassLynx niet meer te gebruiken.`r`n`r`nDoorgaan?" + $extra),
            'MassLynx verwijderen', 'YesNo', 'Warning')
        if ($answer -eq 'Yes') { & $script:RunBlock 'Remove' }
    })

    $btnLog.Add_Click({
        if ($script:LogFile -and (Test-Path $script:LogFile)) { Start-Process notepad.exe $script:LogFile }
        else { Start-Process explorer.exe (Split-Path -Parent (Join-Path $LogDir 'x')) }
    })

    $btnClose.Add_Click({ $script:Form.Close() })

    return $form
}

# ================================================================ hoofdprogramma
function Start-Main {
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
    $script:LogFile = Join-Path $LogDir ('MassLynx-verwijderen-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
    Set-Content -Path $script:LogFile -Value ('MassLynx verwijderen gestart op ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + ' op ' + $env:COMPUTERNAME) -Encoding UTF8

    $script:DoBackup = $true
    if ($NoBackup) { $script:DoBackup = $false }

    if ($Action -eq 'Gui') {
        if (-not (Test-Admin)) {
            $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
                         '-File', ('"' + $PSCommandPath + '"'), '-Action', 'Gui')
            try { Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs; return } catch { }
        }
        Hide-OwnConsole -Always
        $form = Initialize-Gui
        [void]$form.ShowDialog()
        return
    }

    if (-not (Test-Admin)) {
        Write-Host 'Dit script moet als administrator draaien.' -ForegroundColor Red
        exit 2
    }

    Write-Log ('MassLynx verwijderen versie ' + $script:Version + ' - actie: ' + $Action)
    Write-Log ('Computer: ' + $env:COMPUTERNAME + '   logbestand: ' + $script:LogFile)

    $fail = $false
    try {
        switch ($Action) {
            'Check'  { Invoke-ActionCheck }
            'Remove' { $r = Invoke-ActionRemove; if (@($r | Where-Object { -not $_.Ok }).Count -gt 0) { $fail = $true } }
        }
    } catch {
        $fail = $true
        Write-ErrorDetail -ErrorRecord $_
    }

    Write-Log ''
    if ($fail) { Write-Log 'RESULTAAT: NIET IN ORDE - zie de regels met XX hierboven' 'FOUT'; exit 1 }
    Write-Log 'RESULTAAT: IN ORDE' 'OK'
    exit 0
}

Start-Main

<#
.SYNOPSIS
    Starter for Remove-MassLynx: stages the tool to a machine-wide folder
    and starts it elevated.

.DESCRIPTION
    Why this extra step exists (managed PCs):

      * the tool needs administrator rights, and on a managed PC that is a
        DIFFERENT account than the one the operator is logged in as;
      * the elevated process therefore runs as another user, and that user often
        cannot read the operator's profile (own folder, OneDrive, a network home
        drive, a mapped drive letter);
      * result: the tool starts and dies, or cannot find its own files.

    This starter runs as the logged-in user, copies Remove-MassLynx.ps1 to a
    folder that every account can read (C:\ProgramData\...), checks it, and then
    starts it elevated with -Verb RunAs (Windows asks for the administrator
    account).

    Run it:
      powershell -NoProfile -ExecutionPolicy Bypass -File .\Start-MassLynxRemoval.ps1
      powershell ... -Action Check        (read-only check, still elevated)
      powershell ... -Action Remove       (removal without a window)
      powershell ... -StageOnly           (only copy, do not start)
      powershell ... -From D:\x\MassLynx-Onderhoud.ps1

    Without -From and without a copy next to this file, it downloads the script
    from the repository using $env:GH_TOKEN (read-only "Contents" token).
#>
[CmdletBinding()]
param(
    [string] $Repo   = 'MRM15R/masslynx-maintenance',
    [string] $Branch = 'main',
    [string] $From,
    [ValidateSet('Gui','Check','Remove','Reinstall')]
    [string] $Action = 'Gui',
    [string] $Media,
    [switch] $StageOnly,
    [switch] $NoElevate
)

$ErrorActionPreference = 'Stop'
$payload = 'Remove-MassLynx.ps1'
$script:Staged = $null

# Map van deze starter. Buiten een functie vastleggen: binnen een functie wijst
# $MyInvocation naar die functie en niet naar het script (Path is dan leeg).
if ($PSCommandPath) { $script:Here = Split-Path -Parent $PSCommandPath }
else { $script:Here = (Get-Location).Path }

function Say {
    param([string]$Text, [string]$Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
}

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Hide-OwnConsole {
    try {
        if ((Get-Process -Id $PID).MainWindowTitle -ne 'MassLynx verwijderen') { return }
        if (-not ('ConsoleVenster2' -as [type])) {
            Add-Type -Namespace W -Name ConsoleVenster2 -MemberDefinition @'
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
        }
        $h = [W.ConsoleVenster2]::GetConsoleWindow()
        if ($h -ne [IntPtr]::Zero) { [void][W.ConsoleVenster2]::ShowWindow($h, 0) }
    } catch { }
}

function Get-SourceFile {
    # 1. expliciet meegegeven
    if ($From) {
        if (Test-Path -LiteralPath $From) { return (Resolve-Path -LiteralPath $From).Path }
        throw "Bestand niet gevonden: $From"
    }
    # 2. naast deze starter
    $naast = Join-Path $script:Here $payload
    if (Test-Path -LiteralPath $naast) { return $naast }
    # 3. uit de repository (met token als die er is)
    $url = "https://raw.githubusercontent.com/$Repo/$Branch/$payload"
    $tmp = Join-Path $env:TEMP ("ml-" + [Guid]::NewGuid().ToString('N') + ".ps1")
    Say "Downloaden uit $Repo ..."
    try {
        if ($env:GH_TOKEN) {
            Invoke-WebRequest -Uri $url -Headers @{ Authorization = "Bearer $env:GH_TOKEN" } -OutFile $tmp -UseBasicParsing -ErrorAction Stop
        } else {
            Invoke-WebRequest -Uri $url -OutFile $tmp -UseBasicParsing -ErrorAction Stop
        }
        return $tmp
    } catch {
        throw ("Kon $payload niet downloaden uit $Repo ($($_.Exception.Message)). " +
               "Zet het bestand naast deze starter, of geef -From <pad>, of zet `$env:GH_TOKEN " +
               "voor een besloten repository.")
    }
}

function Get-StageFolder {
    $kandidaten = @(
        (Join-Path $env:ProgramData 'MassLynxOnderhoud\app'),
        (Join-Path $env:PUBLIC 'MassLynxOnderhoud\app'),
        (Join-Path $env:LOCALAPPDATA 'MassLynxOnderhoud\app')
    )
    foreach ($k in $kandidaten) {
        try {
            New-Item -ItemType Directory -Force -Path $k | Out-Null
            $probe = Join-Path $k ('.probe-' + [Guid]::NewGuid().ToString('N'))
            Set-Content -Path $probe -Value 'x' -Encoding ASCII -ErrorAction Stop
            Remove-Item $probe -Force -ErrorAction SilentlyContinue
            return $k
        } catch { }
    }
    throw "Geen schrijfbare map gevonden om het script neer te zetten (ProgramData, Public en LocalAppData geweigerd)."
}

function Test-Payload {
    param([string]$Path, [string]$StageDir)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Het script staat niet op $Path" }
    $size = (Get-Item -LiteralPath $Path).Length
    if ($size -lt 20000) { throw "Het script is te klein ($size bytes) - waarschijnlijk niet goed gedownload." }

    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) { throw ("Het script bevat een syntaxfout: " + $errors[0].Message) }

    $versie = 'onbekend'
    $regel = Select-String -LiteralPath $Path -Pattern "script:Version\s*=\s*'([^']+)'" | Select-Object -First 1
    if ($regel) { $versie = $regel.Matches[0].Groups[1].Value }
    return $versie
}

try {
    Hide-OwnConsole

    Say ''
    Say 'MassLynx verwijderen - starter' 'Cyan'
    Say ('  computer : ' + $env:COMPUTERNAME)
    Say ('  account  : ' + [Security.Principal.WindowsIdentity]::GetCurrent().Name)
    Say ('  admin    : ' + (Test-Admin))
    Say ''

    $bron = Get-SourceFile
    Say ("Bron    : " + $bron) 'DarkGray'

    $doelMap = Get-StageFolder
    $script:Staged = Join-Path $doelMap $payload
    Copy-Item -LiteralPath $bron -Destination $script:Staged -Force
    Say ("Kopie   : " + $script:Staged) 'DarkGray'

    $versie = Test-Payload -Path $script:Staged -StageDir $doelMap
    Say ("Versie  : " + $versie + "  (controle ok)") 'Green'

    if ($StageOnly) {
        Say ''
        Say 'Alleen gekopieerd (-StageOnly). Klaar.' 'Green'
        exit 0
    }

    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
                 '-File', ('"' + $script:Staged + '"'), '-Action', $Action)
    if ($Media) { $argList += @('-Media', ('"' + $Media + '"')) }

    if ((Test-Admin) -or $NoElevate) {
        # al verhoogd (of bewust niet): gewoon starten
        Say ''
        Say 'Starten...' 'Cyan'
        & powershell.exe @argList
        exit $LASTEXITCODE
    }

    Say ''
    Say 'Windows vraagt nu om een account met beheerdersrechten.' 'Yellow'
    Say 'Dat mag een ander account zijn dan waarop je nu bent ingelogd.' 'Yellow'
    Say ''
    try {
        $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -PassThru
        Say ("Gestart met beheerdersrechten (proces " + $p.Id + ").") 'Green'
        exit 0
    } catch {
        Say ''
        Say 'Niet gestart: de beheerdersvraag is geannuleerd of niet toegestaan.' 'Red'
        Say ('Melding: ' + $_.Exception.Message) 'Red'
        Say ''
        Say 'Wat je kunt doen:' 'Yellow'
        Say '  - het opnieuw proberen en bij de Windows-vraag het admin-account invullen;'
        Say '  - of IT vragen het script een keer met dat account te starten.'
        exit 2
    }
} catch {
    Say ''
    Say ('FOUT: ' + $_.Exception.Message) 'Red'
    exit 1
}

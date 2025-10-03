# ===== Intune-Upload für 7-Zip (Commands/Detection/Icon) =====
# Struktur (Geschwister-Ordner):
#   7-Zip\
#     ├─ Source\                (wird verpackt; Wrapper werden hier erzeugt)
#     ├─ Commands\              (install.txt / uninstall.txt)
#     ├─ Detection\             (detection.ps1)
#     ├─ Icon\                  (*.png – erstes PNG wird verwendet)
#     └─ Output\                (wird vor jedem Build geleert)

# --- Mandant (Client Secret) ---
$TenantId     = "6ac80ced-0455-4575-86c7-777085e7f076"
$ClientId     = "33461a5c-9560-49d6-8f4a-b05c2d36897a"
$ClientSecret = "h8ZHivI5YKBcT6Lb4Kam8IMvPqrZwohukJcRLeE6Vrc="

$ErrorActionPreference = 'Stop'

# --- Module laden (bei Bedarf installieren) ---
Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction SilentlyContinue
Import-Module IntuneWin32App             -ErrorAction SilentlyContinue
if (-not (Get-Module SvRooij.ContentPrep.Cmdlet)) { Install-Module SvRooij.ContentPrep.Cmdlet -Scope CurrentUser -Force; Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction Stop }
if (-not (Get-Module IntuneWin32App))   { Install-Module IntuneWin32App   -Scope CurrentUser -Force -AcceptLicense; Import-Module IntuneWin32App -ErrorAction Stop }

# --- Pfade (relativ zum Skript) ---
$RootFolder   = $PSScriptRoot
$SourceFolder = Join-Path $RootFolder "Source"
$CmdFolder    = Join-Path $RootFolder "Commands"
$DetectFolder = Join-Path $RootFolder "Detection"
$IconFolder   = Join-Path $RootFolder "Icon"
$OutFolder    = Join-Path $RootFolder "Output"

# --- Eingaben prüfen ---
$InstallTxt   = Join-Path $CmdFolder "install.txt"
$UninstallTxt = Join-Path $CmdFolder "uninstall.txt"
$DetectPs1    = Join-Path $DetectFolder "detection.ps1"   # muss .ps1 sein

foreach ($f in @($InstallTxt,$UninstallTxt,$DetectPs1)) {
  if (-not (Test-Path $f)) { throw "Fehlt: $f" }
}

# --- Icon (optional) ---
$IconPath = Get-ChildItem -Path $IconFolder -Filter *.png -File -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName -ErrorAction SilentlyContinue
$IconB64  = $null
if ($IconPath) { $IconB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($IconPath)) }

# --- Wrapper in Source erzeugen (Befehle werden eingebettet) ---
$InstallLine   = (Get-Content -LiteralPath $InstallTxt   -Raw).Trim()
$UninstallLine = (Get-Content -LiteralPath $UninstallTxt -Raw).Trim()
if ([string]::IsNullOrWhiteSpace($InstallLine))   { throw "install.txt ist leer." }
if ([string]::IsNullOrWhiteSpace($UninstallLine)) { throw "uninstall.txt ist leer." }

$InstallPs1Path   = Join-Path $SourceFolder "Install.ps1"
$UninstallPs1Path = Join-Path $SourceFolder "Uninstall.ps1"

@"
`$cmd = @'
$InstallLine
'@
`$proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `$cmd" -Wait -PassThru
exit `$proc.ExitCode
"@ | Set-Content -Path $InstallPs1Path -Encoding UTF8

@"
`$cmd = @'
$UninstallLine
'@
`$proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `$cmd" -Wait -PassThru
exit `$proc.ExitCode
"@ | Set-Content -Path $UninstallPs1Path -Encoding UTF8

# --- Output frisch anlegen (nicht unterhalb von Source!) ---
if (Test-Path $OutFolder) { Remove-Item $OutFolder -Recurse -Force }
New-Item -ItemType Directory -Path $OutFolder | Out-Null

# --- Verbindung zu Intune (App-Only) ---
Connect-MSIntuneGraph -TenantID $TenantId -ClientId $ClientId -ClientSecret $ClientSecret | Out-Null

# --- Paket bauen (dein Modul: -SourcePath/-DestinationPath; kein -Force) ---
$SetupFile = "Install.ps1"
New-IntuneWinPackage -SourcePath $SourceFolder -SetupFile $SetupFile -DestinationPath $OutFolder

# --- .intunewin ermitteln ---
$intuneWin = Get-ChildItem -Path $OutFolder -Recurse -Filter *.intunewin |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1 -ExpandProperty FullName
if (-not $intuneWin) { throw ".intunewin wurde nicht erstellt in $OutFolder." }

# --- Install-/Uninstall-Kommandos für Intune ---
$InstallCmd   = 'powershell.exe -ExecutionPolicy Bypass -File Install.ps1'
$UninstallCmd = 'powershell.exe -ExecutionPolicy Bypass -File Uninstall.ps1'

# --- Detection aus Detection\detection.ps1 (als Script-Detection) ---
$DetectionRule = New-IntuneWin32AppDetectionRuleScript `
  -ScriptFile $DetectPs1 `
  -EnforceSignatureCheck $false `
  -RunAs32Bit $false

# --- App-Metadaten ---
$DisplayName = "7-Zip"
$Description = "Installiert 7-Zip gemäss Commands\install.txt / uninstall.txt."
$Publisher   = "7-Zip"
$Version     = "23.x"

# --- Upload nach Intune ---
$params = @{
  FilePath             = $intuneWin
  DisplayName          = $DisplayName
  Description          = $Description
  Publisher            = $Publisher
  AppVersion           = $Version
  InstallCommandLine   = $InstallCmd
  UninstallCommandLine = $UninstallCmd
  InstallExperience    = 'system'
  RestartBehavior      = 'suppress'
  DetectionRule        = @($DetectionRule)
}
if ($IconB64) { $params.Icon = $IconB64 }

Add-IntuneWin32App @params

Write-Host "✔ Fertig. Hochgeladen: $intuneWin"
if ($IconB64) { Write-Host "Icon gesetzt: $IconPath" }

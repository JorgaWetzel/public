# ===== Intune-Upload: PS1 direkt aus Commands übernehmen =====
# Erwartete Struktur:
#   Commands\Install.ps1, Commands\Uninstall.ps1, Commands\Detection.ps1
#   Source\  (Ziel für die zwei PS1 vor dem Packen)
#   Icon\*.png (optional)
#   Output\  (wird geleert)

# --- Mandant (Client Secret) ---
$TenantId     = "6ac80ced-0455-4575-86c7-777085e7f076"
$ClientId     = "33461a5c-9560-49d6-8f4a-b05c2d36897a"
$ClientSecret = "h8ZHivI5YKBcT6Lb4Kam8IMvPqrZwohukJcRLeE6Vrc="

$ErrorActionPreference = 'Stop'

# --- Module ---
Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction SilentlyContinue
Import-Module IntuneWin32App             -ErrorAction SilentlyContinue
if (-not (Get-Module SvRooij.ContentPrep.Cmdlet)) { Install-Module SvRooij.ContentPrep.Cmdlet -Scope CurrentUser -Force; Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction Stop }
if (-not (Get-Module IntuneWin32App))   { Install-Module IntuneWin32App   -Scope CurrentUser -Force -AcceptLicense; Import-Module IntuneWin32App -ErrorAction Stop }

# --- Pfade ---
$RootFolder    = $PSScriptRoot
$SourceFolder  = Join-Path $RootFolder "Source"
$CmdFolder     = Join-Path $RootFolder "Commands"
$IconFolder    = Join-Path $RootFolder "Icon"
$OutFolder     = Join-Path $RootFolder "Output"

$CmdInstall    = Join-Path $CmdFolder "Install.ps1"
$CmdUninstall  = Join-Path $CmdFolder "Uninstall.ps1"
$CmdDetection  = Join-Path $CmdFolder "Detection.ps1"

# Eingaben prüfen
foreach ($f in @($CmdInstall,$CmdUninstall,$CmdDetection)) {
  if (-not (Test-Path $f)) { throw "Fehlt: $f" }
}

# Icon (optional)
$IconPath = Get-ChildItem -Path $IconFolder -Filter *.png -File -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName -ErrorAction SilentlyContinue
$IconB64 = $null
if ($IconPath) { $IconB64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($IconPath)) }

# Output frisch
if (Test-Path $OutFolder) { Remove-Item $OutFolder -Recurse -Force }
New-Item -ItemType Directory -Path $OutFolder | Out-Null

# Install/Uninstall-PS1 aus Commands 1:1 nach Source kopieren
Copy-Item -Path $CmdInstall   -Destination (Join-Path $SourceFolder "Install.ps1")   -Force
Copy-Item -Path $CmdUninstall -Destination (Join-Path $SourceFolder "Uninstall.ps1") -Force

# Verbindung
Connect-MSIntuneGraph -TenantID $TenantId -ClientId $ClientId -ClientSecret $ClientSecret | Out-Null

# Paket bauen (SetupFile liegt jetzt in Source)
$SetupFile = "Install.ps1"
New-IntuneWinPackage -SourcePath $SourceFolder -SetupFile $SetupFile -DestinationPath $OutFolder

# .intunewin ermitteln
$intuneWin = Get-ChildItem -Path $OutFolder -Recurse -Filter *.intunewin |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1 -ExpandProperty FullName
if (-not $intuneWin) { throw ".intunewin wurde nicht erstellt in $OutFolder." }

# Detection-Rule (nimmt Commands\Detection.ps1)
$DetectionRule = New-IntuneWin32AppDetectionRuleScript `
  -ScriptFile $CmdDetection `
  -EnforceSignatureCheck $false `
  -RunAs32Bit $false

# Install-/Uninstall-Kommandos (führen die PS1 im Paket aus)
$InstallCmd   = 'powershell.exe -ExecutionPolicy Bypass -File Install.ps1'
$UninstallCmd = 'powershell.exe -ExecutionPolicy Bypass -File Uninstall.ps1'

# Metadaten
$DisplayName = "7-Zip"
$Description = "Installiert 7-Zip gemäss Commands\Install.ps1 / Uninstall.ps1."
$Publisher   = "7-Zip"
$Version     = "23.x"

# Upload
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

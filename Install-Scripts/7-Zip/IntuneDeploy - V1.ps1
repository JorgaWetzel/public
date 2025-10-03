# ===== Minimaler Intune-Upload für 7-Zip (PSADT) =====
# Ordnerstruktur:
#  ...\7-Zip\Source\Deploy-Application.ps1
#  ...\7-Zip\Output\   (leer; wird vom Skript neu erstellt)

# Mandanten-Daten (Client Secret)
$TenantId     = "6ac80ced-0455-4575-86c7-777085e7f076"
$ClientId     = "33461a5c-9560-49d6-8f4a-b05c2d36897a"
$ClientSecret = "h8ZHivI5YKBcT6Lb4Kam8IMvPqrZwohukJcRLeE6Vrc="

$ErrorActionPreference = 'Stop'

# Module
Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction SilentlyContinue
Import-Module IntuneWin32App             -ErrorAction SilentlyContinue
if (-not (Get-Module SvRooij.ContentPrep.Cmdlet)) { Install-Module SvRooij.ContentPrep.Cmdlet -Scope CurrentUser -Force; Import-Module SvRooij.ContentPrep.Cmdlet -ErrorAction Stop }
if (-not (Get-Module IntuneWin32App))   { Install-Module IntuneWin32App   -Scope CurrentUser -Force -AcceptLicense; Import-Module IntuneWin32App -ErrorAction Stop }

# Pfade (Source & Output sind Geschwister)
$RootFolder   = $PSScriptRoot
$SourceFolder = Join-Path $RootFolder "Source"
$OutFolder    = Join-Path $RootFolder "Output"
$SetupFile    = "Deploy-Application.ps1"

if (-not (Test-Path (Join-Path $SourceFolder $SetupFile))) {
    throw "Nicht gefunden: $SourceFolder\$SetupFile"
}

# Output frisch anlegen (nicht in Source!)
if (Test-Path $OutFolder) { Remove-Item $OutFolder -Recurse -Force }
New-Item -ItemType Directory -Path $OutFolder | Out-Null

# Verbindung zu Intune (App-Only)
Connect-MSIntuneGraph -TenantID $TenantId -ClientId $ClientId -ClientSecret $ClientSecret | Out-Null

# Paket bauen (DestinationPath darf kein Unterordner von Source sein)
New-IntuneWinPackage -SourcePath $SourceFolder -SetupFile $SetupFile -DestinationPath $OutFolder

# erzeugte .intunewin ermitteln
$intuneWin = Get-ChildItem -Path $OutFolder -Recurse -Filter *.intunewin |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1 -ExpandProperty FullName
if (-not $intuneWin) { throw ".intunewin wurde nicht erstellt in $OutFolder." }

# --- PSADT Install/Uninstall ---
$InstallCmd   = 'powershell.exe -ExecutionPolicy Bypass -File Deploy-Application.ps1 -DeploymentType Install -DeployMode Silent'
$UninstallCmd = 'powershell.exe -ExecutionPolicy Bypass -File Deploy-Application.ps1 -DeploymentType Uninstall -DeployMode Silent'

# --- Detection-Skript (sehr simpel) ---
$DetectFile = Join-Path $OutFolder "Detect-7Zip.ps1"
@'
$paths = @("$env:ProgramFiles\7-Zip\7z.exe","$env:ProgramFiles(x86)\7-Zip\7z.exe")
if ($paths | Where-Object { Test-Path $_ }) { exit 0 } else { exit 1 }
'@ | Set-Content -Path $DetectFile -Encoding UTF8

# --- DetectionRule-Objekt für Add-IntuneWin32App ---
$DetectionRule = New-IntuneWin32AppDetectionRuleScript `
  -ScriptFile $DetectFile `
  -EnforceSignatureCheck $false `
  -RunAs32Bit $false

# --- Metadaten ---
$DisplayName = "7-Zip (PSADT)"
$Description = "Installiert 7-Zip über PSAppDeployToolkit."
$Publisher   = "7-Zip"
$Version     = "23.x"

# --- Upload ---
Add-IntuneWin32App `
  -FilePath $intuneWin `
  -DisplayName $DisplayName `
  -Description $Description `
  -Publisher $Publisher `
  -AppVersion $Version `
  -InstallCommandLine $InstallCmd `
  -UninstallCommandLine $UninstallCmd `
  -InstallExperience system `
  -RestartBehavior suppress `
  -DetectionRule @($DetectionRule)

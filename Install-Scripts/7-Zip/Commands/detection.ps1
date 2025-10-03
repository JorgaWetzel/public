$ok = $false
try {
  $path = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\7-Zip" -Name "Path" -ErrorAction Stop
  if ($path -and (Test-Path (Join-Path $path "7z.exe"))) { $ok = $true }
} catch {}
exit ($(if ($ok) { 0 } else { 1 }))


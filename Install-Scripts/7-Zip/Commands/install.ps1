$p = Start-Process -FilePath 'cmd.exe' `
  -ArgumentList '/c .\7z2301-x64.exe /S' `
  -WorkingDirectory $PSScriptRoot `
  -Wait -PassThru
exit $p.ExitCode

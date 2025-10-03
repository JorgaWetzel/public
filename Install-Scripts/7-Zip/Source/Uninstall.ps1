$p = Start-Process -FilePath 'cmd.exe' `
  -ArgumentList '/c "%ProgramFiles%\7-Zip\Uninstall.exe" /S' `
  -WorkingDirectory $PSScriptRoot `
  -Wait -PassThru
exit $p.ExitCode

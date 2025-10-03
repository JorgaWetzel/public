@{
  DisplayName       = '7-Zip'
  Description       = 'Installiert 7-Zip gemäss Commands\Install.ps1 / Uninstall.ps1.'
  Publisher         = '7-Zip'
  Version           = '23.x'
  InstallExperience = 'system'
  RemoveExistingAssignments = $true

  InstallGroupName = 'DEV-WIN-Standard'
  InstallGroupId   = '11178c46-7248-4a4b-87c9-c5d59f55c74a'  # <- optionaler Joker, wenn gesetzt wird NICHT aufgelöst
  InstallIntent    = 'required'
  InstallNotification  = 'hideAll'   # showAll | showReboot | hideAll
  
  AvailableToAllUsers   = $true
  AvailableNotification = 'hideAll'   # optional: showAll | showReboot | hideAll

  UninstallGroupName = ''
  UninstallGroupId   = ''
  UninstallIntent    = 'uninstall'
}

    # @{ GroupName = 'Self-Service Apps'; Intent = 'available' }
    # @{ GroupId   = '00000000-0000-0000-0000-000000000000'; Intent = 'uninstall' }

Write-Host "CUSTOM OOBE KEYS : Adding keys..."
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name PrivacyConsentStatus -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name ProtectYourPC -Value 3 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name HideEULAPage -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name HideOEMRegistrationScreen -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name HideWirelessSetupInOOBE -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name SkipMachineOOBE -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE -Name SkipUserOOBE -Value 1 -PropertyType DWORD -Force
Write-Host "CUSTOM OOBE KEYS : Done."
Write-Host "CUSTOM OOBE KEYS FOR PRIVACY : Adding keys..."
New-Item -Path HKLM:\Software\Policies\Microsoft\Windows\OOBE -Force
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\OOBE -Name DisablePrivacyExperience -Value 1 -PropertyType DWORD -Force
New-Item -Path HKLM:\Software\Policies\Microsoft\Windows\DataCollection -Force
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\DataCollection -Name AllowTelemetry -Value 3 -PropertyType DWORD -Force
Write-Host "CUSTOM OOBE KEYS FOR PRIVACY : Done."
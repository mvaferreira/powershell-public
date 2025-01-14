Write-Host "CUSTOM KEYS FOR PRIVACY : Adding key DisablePrivacyExperience..."
New-Item -Path HKLM:\Software\Policies\Microsoft\Windows\OOBE -Force
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\OOBE -Name DisablePrivacyExperience -Value 1 -PropertyType DWORD -Force
Write-Host "CUSTOM KEYS FOR PRIVACY : Adding key AllowTelemetry..."
New-Item -Path HKLM:\Software\Policies\Microsoft\Windows\DataCollection -Force
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\DataCollection -Name AllowTelemetry -Value 3 -PropertyType DWORD -Force
Write-Host "CUSTOM KEYS FOR PRIVACY : Done."
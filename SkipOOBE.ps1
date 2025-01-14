Write-Host "CUSTOM OOBE KEYS FOR PRIVACY : Adding keys..."
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\OOBE -Name DisablePrivacyExperience -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path HKLM:\Software\Policies\Microsoft\Windows\DataCollection -Name AllowTelemetry -Value 3 -PropertyType DWORD -Force
Write-Host "CUSTOM OOBE KEYS FOR PRIVACY : Done."
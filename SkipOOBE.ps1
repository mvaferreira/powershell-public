# Define the content of the unattend.xml file
$unattendXmlContent = @"
<unattend>
  <settings pass="oobeSystem">
    <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <InputLocale>en-US</InputLocale>
      <SystemLocale>en-US</SystemLocale>
      <UILanguage>en-US</UILanguage>
      <UserLocale>en-US</UserLocale>
    </component>
    <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
      <OOBE>
        <HideEULAPage>true</HideEULAPage>
        <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
        <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
        <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
        <HideLocalAccountScreen>true</HideLocalAccountScreen>
        <ProtectYourPC>3</ProtectYourPC>
      </OOBE>
    </component>
  </settings>
</unattend>
"@

# Define the path where the unattend.xml file will be saved
$unattendFilePath = "C:\Windows\Panther\unattend.xml"

# Create the directory if it doesn't exist
if (-not (Test-Path -Path (Split-Path -Parent $unattendFilePath))) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $unattendFilePath) -Force
}

# Write the content to the unattend.xml file
Set-Content -Path $unattendFilePath -Value $unattendXmlContent

Write-Output "unattend.xml file has been created at $unattendFilePath"
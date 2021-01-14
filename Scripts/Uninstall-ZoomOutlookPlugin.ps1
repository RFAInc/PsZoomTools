param($LogPath = 'C:\Windows\Temp\ZoomAddinUninstall.log')

# Prerequisite function
(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/tonypags/PsWinAdmin/master/Get-InstalledSoftware.ps1')|iex;

# Get the uninstall string from HKLM
$RegInfo = Get-InstalledSoftware | Where-Object {$.Name -eq 'Zoom Outlook Plugin'}
$rawUninstallString = $RegInfo.UninstallCommand

# Modify the string
$newUninstallString = $rawUninstallString.Replace('/I','/X').Replace('{','"{').Replace('}','}"')
[string]$UninstallString = $newUninstallString + ' /qn /norestart /log ' + $LogPath

Try {
    
    Write-Verbose "Attempting to remove $($RegInfo.Name) v$($RegInfo.Version) with Install Date $($RegInfo.InstallDate)..."
    $Result = Invoke-Expression $UninstallString -ea Stop -ev rem
    Write-Output $Result

    Write-Verbose "A log is available with the MSI results."
    Get-Item $LogPath -ea 0

} Catch {
    Write-Warning "Exception Raised: $($rem.Exception.Message)"
}

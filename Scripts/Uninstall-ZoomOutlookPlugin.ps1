$LogPath = 'C:\Windows\Temp\ZoomAddinUninstall.log'

# Prerequisite function
(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/tonypags/PsWinAdmin/master/Get-InstalledSoftware.ps1')|iex;

# Get the uninstall string from HKLM
$AppSearchTerm = 'Zoom Outlook Plugin'
$RegInfo = Get-InstalledSoftware | Where-Object {$.Name -eq $AppSearchTerm}
$rawUninstallString = $RegInfo.UninstallCommand

# Check to make sure we have something to do
if (-not $rawUninstallString) {
    Write-Warning "There is no $AppSearchTerm to remove. Exiting Script"
} else {

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

}#END: if (-not $rawUninstallString)

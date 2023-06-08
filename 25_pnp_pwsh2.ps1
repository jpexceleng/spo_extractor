# This script writes all information about all files on a sharepoint site
# to file.

$SiteUrl = "https://exceng.sharepoint.com/sites/Extranet/"
$List = "SharedFiles"
$Fields = "Name","GUID"
$WriteFilePath = "res/SharePointData/SharedFiles.txt"

try {
    # connect to sharepoint site
    Write-Host "Connecting to SharePoint..."
    Connect-PnPOnline -Url $SiteUrl -Interactive

    # extract sharepoint data
    Write-Host "Extracting SharePoint data..."
    $data = (Get-PnPListItem -List $List -Fields $Fields).FieldValues
    
    # write sharepoint data to file
    Write-Host "Writing SharePoint data to file..."
    Write-Output $data | Out-File -FilePath $WriteFilePath
}
catch {
    Write-Error "$($_.Exception.Message)"
}
finally {
    Write-Host "Disconnecting from SharePoint..."
    Disconnect-PnPOnline
}
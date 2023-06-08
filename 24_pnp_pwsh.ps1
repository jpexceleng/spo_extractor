# The following demonstrates how PnP.PowerShell module can be used to 
# programmatically create a new SharePoint page with an embedded file.

$SiteUrl="https://exceng.sharepoint.com/sites/KB/"
$PageName="TestPage"

try {

    # connect to sharepoint site
    Connect-PnPOnline -Url $SiteUrl -Interactive
   
    # create new test page
    Add-PnPPage `
        -Name $PageName `
        -Title "This is a test page" `
        -LayoutType Article `
        -CommentsEnabled:$false

    # add new section to test page
    Add-PnPPageSection `
        -Page $PageName `
        -SectionTemplate OneColumn

    # add new page web part to embed a document page section
    Write-Output "Adding web part..."
    Add-PnPPageWebPart `
        -Page $PageName `
        -DefaultWebPartType DocumentEmbed `
        -Section 1 `
        -Column 1

    # Get page web part
    Write-Output "Getting web part..."
    $Page = Get-PnPPage -Identity $PageName
    $WebPart = Get-PnPPageComponent -Page $Page

    # The following will output the JSON properties of the newly created web part
    # $Test = Get-PnPPageComponent $PageName
    # $Test.PropertiesJson | ConvertFrom-Json | ConvertTo-Json

    # modify page web part properties
    $PropertiesJson = Get-Content "res/propertiesjson.json"
    $PropertiesJson = [string]$PropertiesJson
    Write-Output "Setting web part properties..."
    Set-PnPPageWebPart -Page $Page -Identity $WebPart.InstanceId -PropertiesJson $PropertiesJson

    # remove page web part
    # Remove-PnPPageComponent -Page $Page -InstanceId $WebPart.InstanceId

    # publish page
    Write-Output "Publishing..."
    $Page.Publish()
}
catch {
    Write-Error "Something went wrong! $($_.Exception.Message)"
}
finally {
    # disconnect from sharepoint site
    Write-Output "Disconnecting from SharePoint..."
    Disconnect-PnPOnline
}

function Invoke-CatalogItemDetails {

    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [string] $UpdateId
    )

    

        $BaseUri = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$([uri]::EscapeDataString($UpdateId))"
        $Tabs = @("Overview", "LanguageSelection", "PackageDetails", "InstallDetails")
        $Params = @{
            ContentType = "application/x-www-form-urlencoded"
            UseBasicParsing = $true
            ErrorAction = "Stop"
        }
        $HtmlDocs = @{}
        
        
        
        foreach ($tab in $Tabs) {
            try {
                Set-TempSecurityProtocol

                $response = Invoke-WebRequest -Uri $BaseUri+"#"+$tab @Params
                $HtmlTab = [HtmlAgilityPack.HtmlDocument]::new()
                $HtmlTab.LoadHtml($response.RawContent.ToString())
                
                $CouldNotBeFound = $HtmlTab.GetElementbyId("ctl00_catalogBody_thanksNoUpdate")
                if ($null -eq $CouldNotBeFound) {

                    $HtmlDocs[$tab] = $HtmlTab
                }
                else {
                    throw "Could not retrieve details for the Update with ID $UpdateId"  
                }
                
                Set-TempSecurityProtocol -ResetToDefault
            }catch {
                throw $_
            }
        }

        $HtmlDocs

}
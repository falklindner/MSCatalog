class MSCatalogUpdate {
    [string] $Title
    [string] $Products
    [string] $Classification
    [datetime] $LastUpdated
    [string] $Version
    [string] $Size
    [string] $SizeInBytes 
    [string] $Guid
    [string[]] $FileNames

    MSCatalogUpdate() {}

    MSCatalogUpdate($Row, $IncludeFileNames) {
        $Cells = $Row.SelectNodes("td")
        $this.Title = $Cells[1].innerText.Trim()
        $this.Products = $Cells[2].innerText.Trim()
        $this.Classification = $Cells[3].innerText.Trim()
        $this.LastUpdated = (Invoke-ParseDate -DateString $Cells[4].innerText.Trim())
        $this.Version = $Cells[5].innerText.Trim()
        $this.Size = $Cells[6].SelectNodes("span")[0].InnerText
        $this.SizeInBytes = [Int] $Cells[6].SelectNodes("span")[1].InnerText 
        $this.Guid = $Cells[7].SelectNodes("input")[0].Id
        $this.FileNames = if ($IncludeFileNames) {
            $Links = Get-UpdateLinks -Guid $Cells[7].SelectNodes("input")[0].Id
            foreach ($Link in $Links.Matches) {
                $Link.Value.Split('/')[-1]
            }
        }
    }
}


class MSCatalogUpdateWithDetails : MSCatalogUpdate {

    [string] $UpdateId 
    [string] $Description 
    [string] $Architecture 
    [string] $Classification 
    [string] $SupportedProducts 
    [string] $SupportedLanguages 
    [string] $MsrcNumber 
    [string] $MsrcSeverity 
    [string] $KbArticle 
    [string] $MoreInformationUrl 
    [string] $SupportUrl 
    [PSCustomObject[]] $SupersededBy 
    [PSCustomObject[]] $Supersedes 
    [string] $RestartBehavior 
    [string] $UserInput 
    [string] $Exclusive 
    [string] $RequiresNetwork 
    [string] $UninstallNotes 
    [string] $UninstallSteps 

    MSCatalogUpdateWithDetails($Row, $IncludeFileNames) : base($Row, $IncludeFileNames) {
        $HtmlDocs = Invoke-CatalogItemDetails($this.Guid)

        ## Overview Section
        $HtmlDocOverview = $HtmlDocs["Overview"]
        $this.UpdateId = $HtmlDocOverview.GetElementbyId("ScopedViewHandler_UpdateID").InnerText.Trim()
        $this.Description = $HtmlDocOverview.GetElementbyId("ScopedViewHandler_desc").InnerText.Trim()
        $this.Architecture = $HtmlDocOverview.GetElementbyId("archDiv").ChildNodes[2].InnerText.Trim()
        $this.Classification = $HtmlDocOverview.GetElementbyId("classificationDiv").ChildNodes[2].InnerText.Trim()
        $this.SupportedProducts = $HtmlDocOverview.GetElementbyId("productsDiv").ChildNodes[2].InnerText.Trim() -replace '\s(?=\s|,)','' #removes whitespaces followed by whitspaces or commas.
        $this.SupportedLanguages = $HtmlDocOverview.GetElementbyId("languagesDiv").ChildNodes[2].InnerText.Trim() -replace '\s(?=\s|,)','' 
        $this.MsrcNumber = $HtmlDocOverview.GetElementbyId("securityBullitenDiv").ChildNodes[2].InnerText.Trim()
        $this.MsrcSeverity = $HtmlDocOverview.GetElementbyId("ScopedViewHandler_msrcSeverity").InnerText.Trim()
        $this.KbArticle = $HtmlDocOverview.GetElementbyId("kbDiv").ChildNodes[2].InnerText.Trim()
        $this.MoreInformationUrl = $HtmlDocOverview.GetElementbyId("moreInfoDiv").SelectNodes("div").InnerText.Trim()
        $this.SupportUrl =  $HtmlDocOverview.GetElementbyId("suportUrlDiv").SelectNodes("div").InnerText.Trim()

        ## Package Details Section
        $HtmlDocDetails = $HtmlDocs["PackageDetails"]
        $SupersededByBox = $HtmlDocDetails.GetElementbyId("supersededbyInfo").SelectNodes("div")
        $this.SupersededBy = [MSCatalogItemReferenceList]::new($SupersededByBox).ItemReferenceList         
        $SupersedesBox = $HtmlDocDetails.GetElementbyId("supersedesInfo").SelectNodes("div")
        $this.Supersedes = [MSCatalogItemReferenceList]::new($SupersedesBox).ItemReferenceList

        ## Install Details Section
        $HtmlDocInstall = $HtmlDocs["InstallDetails"]
        $this.RestartBehavior = $HtmlDocInstall.GetElementbyId("ScopedViewHandler_rebootBehavior").InnerText.Trim()
        $this.UserInput = $HtmlDocInstall.GetElementbyId("ScopedViewHandler_userInput").InnerText.Trim()
        $this.Exclusive = $HtmlDocInstall.GetElementbyId("ScopedViewHandler_installationImpact").InnerText.Trim()
        $this.RequiresNetwork = $HtmlDocInstall.GetElementbyId("ScopedViewHandler_connectivity").InnerText.Trim()
        $this.UninstallNotes = $HtmlDocInstall.GetElementbyId("uninstallNotesDiv").SelectNodes("div").InnerText.Trim()
        $this.UninstallSteps = $HtmlDocInstall.GetElementbyId("uninstallStepsDiv").SelectNodes("div").InnerText.Trim()
    }
}

class MSCatalogItemReferenceList {
    [PSCustomObject[]] $ItemReferenceList = @()
    
    MSCatalogItemReferenceList($InfoBox) {

        foreach ($node in $InfoBox) {
            $ItemReference = [PSCustomObject]@{
                Name      = $node.InnerText.Trim()
                KbArticle = [string]::Empty
                UpdateID  = [string]::Empty
            }

            $ItemReference.KbArticle = if ($ItemReference.Name -match 'KB\d+'){
                $Matches.0
            }
            $ItemReference.UpdateID = if ($node.SelectNodes("a").Count -eq 1)  {
                
                $href = $node.SelectNodes("a").GetAttributeValue("href", [string]::Empty)
                if ($href.contains("updateid")) {
                    ($href -split ("updateid="))[1].Trim()
                }
            }
        $this.ItemReferenceList += $ItemReference
        }
    }
}

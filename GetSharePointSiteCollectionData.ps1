<#
#Setup an Azure App with SharePoint Permissions - Sites.FullControl.All 
#PnP.PowerShell Version: 1.6.46
#>
function Template{
    Param(
        $template
    )

    $templatehash = @{
        "EHS#1" = 'Initial Site';
        "Group#0" = 'Office 365 Group Team Site';
        "STS#3" = 'Modern Team Site';
        "STS#0" = 'Classic Teams Site';
        "APPCATALOG#0" = 'App Catalog Site';
        "TEAMCHANNEL#0" = 'Private Channel Site';
        "SITEPAGEPUBLISHING#0" = 'Communication Site';
        "RedirectSite#0" = 'Redirect Url Links';
        "PWA#0" = 'Project Web App Site'
    }

    if($templatehash.ContainsKey($template)){
        return $templatehash[$template]
    }else{
        return $template
    }

}
function SiteCollectionAdmins{
    Param(
        $sitecollectionadmin
    )

    $tempstring = ""

    foreach($s in $sitecollectionadmin){
        $tempstring += $s.Title+";"
    }

    return $tempstring.Trim(";")
}
function CheckedOutFiles{
    $lists = Get-PnPList
    $checkedoutfilescount = 0
    $checkedoutinlibrarycount = 0
    foreach($list in $lists){
        if($list.BaseTemplate -eq "101"){

            $checkedoutitems = $list.GetCheckedOutFiles()
            $context = Get-PnPContext
            $context.Load($checkedoutitems)
            Invoke-PnPQuery
            $checkedoutfilescount += $checkedoutitems.Count

            $ListItems = Get-PnPListItem -List $list.Title            
            Foreach($ListItem in $ListItems){
                if ($null -ne $ListItem.FieldValues.CheckoutUser.LookupValue){
                    $checkedoutinlibrarycount++
                }
            }
        }
    }

    $global:checkedoutfiles += $checkedoutinlibrarycount + $checkedoutfilescount

    return $checkedoutfilescount,$checkedoutinlibrarycount
}
function SubSites{
    $SubSites = Get-PnPSubWebs -Includes LastItemUserModifiedDate, Url
    foreach($SubSite in $SubSites){
        $connection = Connect-PnPOnline -Url $SubSite.Url -Tenant <tenant>.onmicrosoft.com -ClientId <clientid> -Thumbprint <certificatethumbprint>
        
        $serverrelativeurl = Get-PnPWeb | Select -ExpandProperty ServerRelativeUrl

        $output = Invoke-PnPSPRestMethod -Url "/_api/web/getFolderByServerRelativeUrl('$serverrelativeurl')?`$select=StorageMetrics&`$expand=StorageMetrics"

        $StorageMetrics = ($Output.TotalSize/1024)

        $subsite = $True

        $Web = Get-PnPWeb

        $ParentSiteUrl = Get-PnPSite | Select -ExpandProperty Url

        CaptureData $Web $SubSite $ParentSiteUrl $StorageMetrics



    }
}
function CaptureData{


    Param(
        $Site,
        $SubSite,
        $ParentSiteUrl,
        $StorageMetrics
    )

    if($SubSite){
        $Template="Same as Parent Site Collection"
        $Size = $StorageMetrics
        $ParentSiteUrl = $ParentSiteUrl

    }else{
        #Get Storage Usage
        $Size = $Site.StorageUsageCurrent
        #Get Template
        $Template = Template $Site.Template
        #Get Status
        $Status = $Site.Status
        $ParentSiteUrl = "This site is Parent site collection"
    }
    #Get Title
    $Title = $Site.Title
    #Get Url
    $Url = $Site.Url
    #Get Subsite
    $Subsites = Get-PnPSubWebs -Recurse
    $SubSitesCount = $SubSites.count
    #Get Site Collection Administrators
    $administrators = Get-PnPSiteCollectionAdmin | Select Title
    $AdministratorsString = SitecollectionAdmins $administrators
    #Get Last Modified by user date
    $LastModifiedDate = Get-PnPWeb -Includes LastItemUserModifiedDate | Select -ExpandProperty LastItemUserModifiedDate
    #Get Never Checked In files
    $FilesNeverCheckedIn, $FilesCheckedOutInLibrary = CheckedOutFiles

    $tempdata = [PSCustomObject]@{
        Title = $Title
        Url = $Url
        Size = $Size
        Template = $Template
        Status = $Status
        SubSitesCount = $SubSitesCount
        "SiteCollection Administrators" = $AdministratorsString
        "Last User Modified Date" = $LastModifiedDate
        "Files were never checked in"= $FilesNeverCheckedIn
        "Files Checked Out In Library"= $FilesCheckedOutInLibrary
        ParentSite = $ParentSiteUrl
    }

    $global:hashitems.Add($Url,$tempdata)

    if($SubSitesCount -gt 0){
        SubSites
    }

}
function main{
    $mainconnection = Connect-PnPOnline -Url https://<tenant>-admin.sharepoint.com -Tenant <tenant>.onmicrosoft.com -ClientId <clientid> -Thumbprint <certificatethumbprint>

    #Get All Site Collections 
    $Sites = Get-PnPTenantSite -Detailed

    #intialize a hashtable
    $global:hashitems = @{}
    $global:checkedoutfiles =0

    #Loop through each Site Collection
    ForEach ($Site in $Sites) {

        $temphash = @{}

        "Processing data for "+ $Site.Url
        $Connection = Connect-PnPOnline -Url $Site.Url -Tenant <tenant>.onmicrosoft.com -ClientId <clientid> -Thumbprint <certificatethumbprint>
        
        $SubSite=$False
        if($Site.Template -ne "RedirectSite#0"){
            CaptureData $Site $SubSite
        }
        

    }
    Write-Host "####Analytics####"
    Write-Host "Total sites data collected for: " $global:hashitems.Count
    Write-Host "Total Checked Out files: "$global:checkedoutfiles
    Write-Host "####https://github.com/sairao77/M365PowerShellScripts####"
}
main
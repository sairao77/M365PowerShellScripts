<#
Name: OneDriveReporting
CreatedBy: Sai Gutta
CreatedOn: 07/18/2021
UpdatedOn: 10/09/2021
PnPPSVersion: 1.7.64-nightly
Contact: sairao77.github.io
#>

#Gathering Pre-Requisite variables

#Provide tenant name of your Microsoft365 env
#TODO: Enter Values
#example: contoso.onmicrosoft.com
$global:tenant = ""

#Provide sharepoint admin url of your Microsoft365 env
#TODO: Enter Values
#example: https://contoso-admin.sharepoint.com
$global:SPadminurl = ""

#Provide azure clientid for app to authenticate
#TODO: Enter Values
#example: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
$global:clientid = ""

#Provide azure Thumbprint for app to authenticate
#TODO: Enter Values
#example: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
$global:thumbprint = ""


#Built in lists in OneDrive
$global:builtinlists = @(
    "Style Library",
    "Social",
    "Form Templates"
)




#Combine Secondary SiteCollection Admins to a string
function SiteCollectionAdmins{
    Param(
        $sitecollectionadmin,
        $owner
    )

    $tempstring = ""

    foreach($s in $sitecollectionadmin){
        if($s.Email -ne $owner){
            $tempstring += $s.Title+";"
        }
        
    }

    return $tempstring.Trim(";")
}

function filecount{
    $totalfilecount = 0
    $lists = Get-PnPList -Includes BaseType, Hidden
    foreach($list in $lists){
        if(($global:builtinlists -NotContains $list.title) -and ($list.BaseType -eq "DocumentLibrary") -and ($list.Hidden -eq $false)){
            $totalfilecount += (Get-PnPlistItem -List $list).Count
        }
    }
    return $totalfilecount
}

function customlists{
    $totallistcount = 0
    $lists = Get-PnPList -Includes BaseType, Hidden
    foreach($list in $lists){
        if(($global:builtinlists -NotContains $list.title) -and ($list.Hidden -eq $false)){
            $totallistcount++
        }
    }
    return $totallistcount
}


#Connect to PnP SharePoint
$mainconnection = Connect-PnPOnline -Tenant $global:tenant -Url $global:SPadminurl -ReturnConnection -ClientId $global:clientid -Thumbprint $global:thumbprint


#Get All OneDrive SiteCollections in a tenant, capture only certain metadata
$OneDriveSites = Get-PnPTenantSite -IncludeOneDriveSites -Template "SPSPERS#10" -Detailed -Connection $mainconnection | `
Select-Object Title, StorageUsageCurrent, Url, Owner, Status



$completedetails = @()
foreach($ODS in $OneDriveSites){

        #Connect to specific OneDrive to gather more details
        $specificweb = Connect-PnPOnline -ReturnConnection -Url $ODS.Url -Tenant $global:tenant -ClientId $global:clientid -Thumbprint $global:thumbprint
        $LastModifiedDate = Get-PnPWeb -connection $specificweb -Includes LastItemUserModifiedDate | Select -ExpandProperty LastItemUserModifiedDate
        $allscadmins = Get-PnPSiteCollectionAdmin -Connection $specificweb
        $secondaryadministrators = SiteCollectionAdmins $allscadmins $ODS.Owner

        #Total file count and item count in OneDrive
        $filecount = filecount
        #List count that are created by users
        $customlistcount = customlists

        $temp = [PSCustomObject]@{
                    Title = $ODS.Title
                    Url = $ODS.Url
                    Size = $ODS.StorageUsageCurrent
                    LastModifiedDate = $LastModifiedDate
                    PrimaryAdministrator = $ODS.Owner
                    SecondaryAdministrators = $secondaryadministrators
                    Status = $ODS.Status
                    CustomListCount = $customlistcount
                    FileCount = $filecount

        }
        $completedetails += $temp
}

#Disconnect all the connections
Disconnect-PnPOnline
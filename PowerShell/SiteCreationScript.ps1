. $PSScriptRoot/include.ps1

function Getitems() {
    Connect -Url "$tenantURL$SiteCreationURL"
    $GetCreateditemsCAML = @"
    <View>
    <Query>
        <Where>
                <Eq>
                    <FieldRef Name='Status' />
                    <Value Type='Choice'>Pending</Value>
                </Eq>
        </Where>
    </Query>
    <ViewFields>
        <FieldRef Name='ID' />
    </ViewFields>
    <QueryOptions />
    </View>
"@
    return @(Get-PnPListItem -List $NewSiteCreationRequestList -Query $GetCreateditemsCAML)
}

$variablesSet = CheckEnvironmentalVariables
if ($variablesSet -eq $false) {    
    Write-Output "Missing one of the following environmental variables: 
    TenantURL,PrimayOwnerEmail,AppId,
    AppSecret,SiteCreationURL,StorageQuota
    ResourceSize"
    exit
}

$GetItemsBySPList = Getitems  
if (!$GetItemsBySPList -or ($GetItemsBySPList -ne $null -and `
(0 -eq $GetItemsBySPList.Count))) {
    Write-Output "No site requests detected"
}

foreach ($siteItem in $GetItemsBySPList) {
    Connect -Url "$tenantURL$SiteCreationURL"
    $siteItem = Get-PnPListItem -List $NewSiteCreationRequestList -Id $siteItem.ID #load all fields
    $title = $siteItem["Title"]
    $status = $siteItem["Status"]
    $siteDescription = $siteItem["SiteDescription"]
    $primaryAdministrator = @($siteItem["PrimaryAdmin"] |? {-not [String]::IsNullOrEmpty($_.Email)} | select -ExpandProperty Email)
    $secondaryAdministrator = @($siteItem["SecondaryAdmin"] |? {-not [String]::IsNullOrEmpty($_.Email)} | select -ExpandProperty Email)
    $siteUrl = GetUniqueUrlFromName -title $title 
    $Global:siteEntryId = [int]$siteItem.ID
    
    write-output $siteUrl

    EnsureSite  -title $title `
        -siteUrl $siteUrl `
        -description $siteDescription `
        -siteCollectionAdmin $PrimaryAdministrator

    if ($? -eq $true) {
        Set-SecondaryAdmin -siteUrl $siteUrl -siteCollectionAdmin $SecondaryAdministrator
        if ($Status -eq 'Pending') {
            SetSiteUrl -siteItem $siteItem -siteUrl $siteUrl -title $title
            UpdateStatus -id $siteItem["ID"] -status 'Created'
        }
    }
}

Disconnect-PnPOnline
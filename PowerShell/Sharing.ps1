. $PSScriptRoot/include.ps1

function Getitems() {
    Connect -Url "$tenantURL$SiteCreationURL"
    $GetCreateditemsCAML = @"
    <View>
    <Query>
        <Where>
                <Eq>
                    <FieldRef Name='ExternalSharingStatus' />
                    <Value Type='Choice'>Approved</Value>
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
if (!$GetItemsBySPList -or `
($GetItemsBySPList -ne $null -and (0 -eq $GetItemsBySPList.Count))) {
    Write-Output "No site requests detected"
}

foreach ($siteItem in $GetItemsBySPList) {
    Connect -Url "$tenantURL$SiteCreationURL"
    $siteItem = Get-PnPListItem -List $NewSiteCreationRequestList -Id $siteItem.ID #load all fields
    $ExternalSharingStatus = $siteItem["ExternalSharingStatus"]
    if ($siteItem["SiteUrl"] -eq $null) {
        Write-Output "Site Url not found"
        exit
    }
    else {
        $siteUrl = $siteItem["SiteUrl"].Url
        Set-ExternalUserSharingOnly -siteUrl $siteUrl
        if ($? -eq $true) {
            if ($ExternalSharingStatus -eq 'Approved') {
                UpdateExternalStatus -id $siteItem["ID"] -status 'Enable'
            }
        }
    }
}

Disconnect-PnPOnline
$ProgressPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
Import-Module $PSScriptRoot\SharePointPnPPowerShellOnline\SharePointPnPPowerShellOnline.psd1 -ErrorAction SilentlyContinue

function Connect([string]$Url) {    
    if ($Url -eq $Global:lastContextUrl) {
        return
    }
    if ($appId -ne $null -and $appSecret -ne $null) {
        Connect-PnPOnline -Url $Url -AppId $appId -AppSecret $appSecret
    }
    else {
        Connect-PnPOnline -Url $Url
    }
    $Global:lastContextUrl = $Url
}

function UpdateStatus($id, $status) {
    Connect -Url "$tenantURL$SiteCreationURL"
    Set-PnPListItem -List $NewSiteCreationRequestList -Identity $id `
     -Values @{"Status" = $status} -ErrorAction SilentlyContinue >$null 2>&1
}

function UpdateExternalStatus($id, $status) {
    Connect -Url "$tenantURL$SiteCreationURL"
    Set-PnPListItem -List $NewSiteCreationRequestList -Identity $id `
    -Values @{"ExternalSharingStatus" = $status} -ErrorAction SilentlyContinue >$null 2>&1
}

function EnsureSite {
    Param(
        [string]$title,
        [string]$siteUrl,
        [string]$description = "",
        [string]$siteCollectionAdmin,
        [string]$SiteTemplate
    )

    $TotalStorateInBytes = ($StorageQuota * 1024)
    $EightyPercentage = ($TotalStorateInBytes * (0.8))
    
    #Connect admin url
    Connect -Url $tenantAdminUrl
    $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue

    if ( $? -eq $false) {
        Write-Output "Site at $siteUrl does not exist - let's create it"
        $parameters = @{
            'Title'                    = $title;
            'Url'                      = $siteUrl ;
            'Owner'                    = $siteCollectionAdmin ;
            'TimeZone'                 = 10 ;
            'Description'              = $description ;
            'ResourceQuota'            = $ResourceQuota ;
            'StorageQuota'             = $TotalStorateInBytes ;
            'StorageQuotaWarningLevel' = $EightyPercentage ;
            'Lcid'                     = 1033 ;
            'RemoveDeletedSite'        = $true ;
            'Wait'                     = $true;
        }

        $parameters['Template'] = 'STS#0'
        $site = New-PnPTenantSite @parameters

        if ($? -eq $false) {
            Write-Output "Site creation fail $siteUrl"
            Write-Error "Something happened"
            UpdateStatus -id $Global:siteEntryId -status 'Error'
            exit
        }
    }
    elseif ($site.Status -ne "Active") {
        Write-Output "Site at $siteUrl already exist"
        while ($true) {
            # Wait for site to be ready
            $site = Get-PnPTenantSite -Url $siteUrl
            if ( $site.Status -eq "Active" ) {
                break;
            }
            Write-Output "Site not ready"
            Start-Sleep -s 20
        }
    }
}

function Set-ExternalUserSharingOnly {
    Param(
        [string]$siteUrl
    )
    Write-Output "External Sharing and Secondary Admin"
    $IsExist = $true
    Connect -Url $tenantAdminUrl
    $siteUrl
    $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
    if ( $? -eq $false) {
        $IsExist = $false
        Write-Output "Site at $siteUrl does not exist"
        exit
    }
    elseif ($site.Status -ne "Active") {
        Write-Output "Site at $siteUrl already exist"
        while ($true) {
            # Wait for site to be ready
            $site = Get-PnPTenantSite -Url $siteUrl
            if ( $site.Status -eq "Active" ) {
                $IsExist = $true
                break;
            }
            Write-Output "Site not ready"
            Start-Sleep -s 20
        }
    }

    if ($IsExist) {
        Write-Output "enable external sharing $siteUrl"
        Set-PnPTenantSite -Url $siteUrl `
            -Sharing ExternalUserSharingOnly -Wait
    }

}


function GetUniqueUrlFromName($title) {
    Connect -Url $tenantAdminUrl
    $cleanName = $title -replace '[^a-z0-9]'
    if ([String]::IsNullOrWhiteSpace($cleanName)) {
        $cleanName = "teams"
    }
    $url = "$tenantUrl/$managedPath/$cleanName"
    $doCheck = $true
    while ($doCheck) {
        Get-PnPTenantSite -Url $url  -ErrorAction SilentlyContinue >$null 2>&1
        if ($? -eq $true) {
            $url += '1'
        }
        else {
            $doCheck = $false
        }
    }
    return $url
}

function CheckEnvironmentalVariables {
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_TenantURL")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_PrimayOwnerEmail")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_AppId")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_AppSecret")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_SiteCreationURL")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_StorageQuota")) {
        return $false
    }
    if (-not [environment]::GetEnvironmentVariable("APPSETTING_ResourceSize")) {
        return $false
    }
}



function SetSiteUrl($siteItem, $siteUrl, $title) {
    Connect -Url "$tenantURL$SiteCreationURL"
    Write-Output "`tSetting site URL to $siteUrl"
    Set-PnPListItem -List $NewSiteCreationRequestList -Identity $siteItem["ID"] `
     -Values @{"SiteUrl" = "$siteUrl, $title"} -ErrorAction SilentlyContinue >$null 2>&1
}


function Set-SecondaryAdmin {
    Param(
        [string]$siteUrl,
        [string]$siteCollectionAdmin
    )
    $IsExist = $true
    Connect -Url $tenantAdminUrl
    
    $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
    if ( $? -eq $false) {
        $IsExist = $false
        Write-Output "Site at $siteUrl does not exist"
    }
    elseif ($site.Status -ne "Active") {
        Write-Output "Site at $siteUrl already exist"
        while ($true) {
            # Wait for site to be ready
            $site = Get-PnPTenantSite -Url $siteUrl
            if ( $site.Status -eq "Active" ) {
                $IsExist = $true
                break;
            }
            Write-Output "Site not ready"
            Start-Sleep -s 20
        }
    }

    if ($IsExist) {
        Write-Output "Set Secondary Admin $siteUrl"
        Set-PnPTenantSite -Url $siteUrl `
            -Owners  @($siteCollectionAdmin) 
        
    }
}


Set-PnPTraceLog -Off
$variablesSet = CheckEnvironmentalVariables
if ( $variablesSet -eq $false) {
    exit
}

$tenantURL = ([environment]::GetEnvironmentVariable("APPSETTING_TenantURL"))
if (!$tenantURL) {
    $tenant = ([environment]::GetEnvironmentVariable("APPSETTING_Tenant"))
    $tenantURL = [string]::format("https://{0}.sharepoint.com", $tenant)
}

$primarySiteCollectionAdmin = ([environment]::GetEnvironmentVariable("APPSETTING_PrimayOwnerEmail"))
$SiteCreationURL = ([environment]::GetEnvironmentVariable("APPSETTING_SiteCreationURL"))
$StorageQuota = [int] ([environment]::GetEnvironmentVariable("APPSETTING_StorageQuota"))
$ResourceQuota = ([environment]::GetEnvironmentVariable("APPSETTING_ResourceSize"))

$NewSiteCreationRequestList = '/Lists/NewSiteCreationRequest'
$managedPath = 'teams' # sites/teams
$Global:lastContextUrl = ''
$Global:siteEntryId = 0

$appId = ([environment]::GetEnvironmentVariable("APPSETTING_AppId"))
if (!$appId) {
    $appId = ([environment]::GetEnvironmentVariable("APPSETTING_ClientId"))
}

$appSecret = ([environment]::GetEnvironmentVariable("APPSETTING_AppSecret"))
if (!$appSecret) {
    $appSecret = ([environment]::GetEnvironmentVariable("APPSETTING_ClientSecret"))
}


$uri = [Uri]$tenantURL
$tenantUrl = $uri.Scheme + "://" + $uri.Host
$tenantAdminUrl = $tenantUrl.Replace(".sharepoint", "-admin.sharepoint")
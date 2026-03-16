<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to assess a Microsoft 365 tenant for migration of Exchange, Teams, SharePoint and OneDrive, 
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Reports on multiple factors of a Microsoft 365 tenant to help with migration preparation. Exports results to Excel

        .DESCRIPTION
        Gathers information using Microsoft Graph API and Exchange Online Management Shell and Exports to CSV

        .PARAMETER ClientID
        Required - Application (Client) ID of the App Registration

        .PARAMETER TenantID
        Required - Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER certificateThumbprint
        Required - Thumbprint of the certificate generated from the prepare-tenantassessment.ps1 script
        
        .PARAMETER IncludeGroupMembership
        Optional - Switch to include group membership in the report

        .PARAMETER IncludeMailboxPermissions
        Optional - Switch to include mailbox permissions in the report

        .PARAMETER IncludeDocumentLibraries
        Optional - Switch to include document libraries in the report

        .PARAMETER IncludeLists
        Optional - Switch to include lists in the report

        .PARAMETER IncludePlans
        Optional - Switch to include Planner plans in the report

        .EXAMPLE
        Perform-TenantAssessment.ps1 -ClientId "12345678-1234-1234-1234-123456789012" -TenantId "12345678-1234-1234-1234-123456789012" -certificateThumbprint "1234567890123456789012345678901234567890" -IncludeGroupMembership -IncludeMailboxPermissions -IncludeDocumentLibraries -IncludeLists -IncludePlans

        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/
        
        For full instructions on how to use this script, please visit the blog posts below:
        https://practical365.com/office-365-migration-plan-assessment/
        https://practical365.com/microsoft-365-tenant-to-tenant-migration-assessment-version-2/

        .VERSION
        2.8 - 2025-12-05
        - If SharePoint usage report cannot be retrieved (e.g., 401 due to missing
          Reports.Read.All) or returns no rows, fall back to Get-MgSite -All and
          build a basic SharePoint Sites dataset so the "SharePoint Sites" tab is
          never empty.

        2.7 - 2025-12-04
        - Added runtime timer: script now records start/end time and prints a
          green summary showing total runtime (hh:mm:ss and minutes) at the end.

        2.6 - 2025-12-04
        - Added timestamp to output Excel file name (TenantAssessment-YYYYMMdd-HHmmss.xlsx)
          so each run creates a unique file and never overwrites previous results.

        2.5 - 2025-12-04
        - Switched Graph usage reports (OneDrive, SharePoint, Mailbox, M365 Apps)
          from Get-MgReport* cmdlets to direct REST calls using the Graph access token
          to avoid 'PercentComplete cannot be greater than 100' errors.

        2.4 - 2025-12-04
        - Worked around Get-MgReport* progress bug ("PercentComplete cannot be greater than 100")
          by temporarily setting $ProgressPreference='SilentlyContinue' around those cmdlets.

        2.3 - 2025-12-04
        - Fixed Connect-MgGraph -AccessToken type mismatch by ensuring the token
          is passed as SecureString (ConvertTo-SecureString when needed).

        2.2 - 2025-12-04
        - Optimized module checking: removed Get-Module -ListAvailable calls and now rely on
          Import-Module with try/catch per module (faster startup).
        - Kept Graph auth via Az.Accounts + Get-AzAccessToken + Connect-MgGraph -AccessToken
          to avoid MSAL ClientCertificateCredential / WithLogging() issues.

        2.1 - 2025-12-04
        - Switched Graph authentication from Connect-MgGraph -Certificate to token-based auth
          using Az.Accounts (Connect-AzAccount + Get-AzAccessToken) and Connect-MgGraph -AccessToken
          to avoid MSAL ClientCertificateCredential / WithLogging() issues.
        - Added Az.Accounts module check and imports.
        - Fixed minor bug in CAHeadings array (missing comma between includeLocations/excludeLocations).
    #>

Param(
    [parameter(Mandatory = $true)]
    $clientId,
    [parameter(Mandatory = $true)]
    $tenantId,
    [parameter(Mandatory = $true)]
    $certificateThumbprint,
    [parameter(Mandatory = $false)]
    [switch]$IncludeGroupMembership = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeMailboxPermissions = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeDocumentLibraries = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludeLists = $false,
    [parameter(Mandatory = $false)]
    [switch]$IncludePlans = $false
)

# === Runtime timer: record start time ===
$ScriptStartTime = Get-Date

function UpdateProgress {
    Write-Progress -Activity "Tenant Assessment in Progress" -Status "Processing Task $ProgressTracker of $($TotalProgressTasks): $ProgressStatus" -PercentComplete (($ProgressTracker / $TotalProgressTasks) * 100)
}

$ProgressTracker    = 1
$TotalProgressTasks = 28
$ProgressStatus     = $null

if ($IncludeGroupMembership)    { $TotalProgressTasks++ }
if ($IncludeMailboxPermissions) { $TotalProgressTasks++ }
if ($IncludeDocumentLibraries)  { $TotalProgressTasks++ }
if ($IncludeLists)              { $TotalProgressTasks++ }
if ($IncludePlans)              { $TotalProgressTasks++ }

$ProgressStatus = "Checking & loading required modules..."
UpdateProgress
$ProgressTracker++

# Faster module handling: try Import-Module once per required module
$requiredModules = @(
    @{
        Name        = 'Microsoft.Graph.Authentication'
        InstallHint = 'Install-Module Microsoft.Graph -Scope CurrentUser'
    },
    @{
        Name        = 'ExchangeOnlineManagement'
        InstallHint = 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    },
    @{
        Name        = 'ImportExcel'
        InstallHint = 'Install-Module ImportExcel -Scope CurrentUser'
    },
    @{
        Name        = 'Az.Accounts'
        InstallHint = 'Install-Module Az.Accounts -Scope CurrentUser'
    }
)

foreach ($mod in $requiredModules) {
    try {
        Import-Module -Name $mod.Name -ErrorAction Stop
    }
    catch {
        Write-Host "Module '$($mod.Name)' is not installed or failed to load." -ForegroundColor Red
        Write-Host "Please install it with:" -ForegroundColor Red
        Write-Host "    $($mod.InstallHint)" -ForegroundColor Yellow
        Pause
        exit 1
    }
}

$ProgressStatus = "Connecting to Microsoft Graph..."
UpdateProgress
$ProgressTracker++

# Connect to Graph using Az.Accounts + AccessToken (avoid ClientCertificateCredential path)
try {
    $CertificatePath = "cert:\CurrentUser\My\$CertificateThumbprint"
    $Certificate     = Get-Item $CertificatePath -ErrorAction Stop

    # Auth as service principal with cert
    Connect-AzAccount -ServicePrincipal `
                      -Tenant $TenantId `
                      -ApplicationId $ClientId `
                      -CertificateThumbprint $certificateThumbprint `
                      -ErrorAction Stop | Out-Null

    # Get Graph token
    $tokenObject = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/" -ErrorAction Stop

    # Keep plain token for REST calls
    $GraphAccessTokenPlain = [string]$tokenObject.Token

    # Ensure token is SecureString – required by Connect-MgGraph -AccessToken
    if ($tokenObject.Token -is [System.Security.SecureString]) {
        $secureToken = $tokenObject.Token
    }
    else {
        $secureToken = ConvertTo-SecureString -String $GraphAccessTokenPlain -AsPlainText -Force
    }

    # Feed token into Graph – bypasses ClientCertificateCredential/MSAL logging issue
    Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop
}
catch {
    Write-Host "Unable to connect to Microsoft Graph." -ForegroundColor Red

    if ($_.Exception.Message -like "*BaseAbstractApplicationBuilder`1.WithLogging*") {
        Write-Host @"
This looks like the known MSAL assembly clash between Microsoft.Graph and other MSAL-based modules.

Workarounds:
  - Run this script in a fresh PowerShell session (no pre-loaded EXO/Az/MSAL.PS), OR
  - Run in Windows PowerShell 5.1, OR
  - Adjust module versions (e.g. older EXO or Graph) if needed.

The script now uses token-based Graph auth via Az.Accounts, so if this still fails,
something else in the session may be injecting an incompatible MSAL assembly.
"@ -ForegroundColor Yellow
    }
    else {
        Write-Host "Raw error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    exit 1
}

$ProgressStatus = "Preparing environment..."
UpdateProgress
$ProgressTracker++

##Report File Name
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$Filename  = "TenantAssessment-$timestamp.xlsx"

##File Location
$FilePath = ".\TenantAssessment"
try {
    if (-not (Test-Path -Path $FilePath)) {
        New-Item -Path $FilePath -ItemType Directory | Out-Null
    }
}
catch {
    Write-Host "Could not create folder at $FilePath - check you have appropriate permissions" -ForegroundColor Red
    exit
}

##Check if cover page is present
$TemplatePath    = "TenantAssessment-Template.xlsx"
$TemplatePresent = Test-Path $TemplatePath

$ProgressStatus = "Getting users..."
UpdateProgress
$ProgressTracker++

##List All Tenant Users
$users = Get-MgUser -All -Property id, userprincipalname, mail, displayname, givenname, surname, licenseAssignmentStates, proxyaddresses, usagelocation, usertype, accountenabled, onPremisesSyncEnabled

$ProgressStatus = "Getting groups..."
UpdateProgress
$ProgressTracker++

##List all Tenant Groups
$Groups = Get-MgGroup -All

$ProgressStatus = "Getting Teams..."
UpdateProgress
$ProgressTracker++

##Get Teams details
$TeamGroups = $Groups | Where-Object { ($_.grouptypes -Contains "unified") -and ($_.additionalproperties.resourceProvisioningOptions -contains "Team") }

$i = 1

foreach ($teamgroup in $TeamGroups) {

    $ProgressStatus = "Processing Team $i of $($Teamgroups.count)..."
    UpdateProgress
    $i++

    [array]$Teamchannels          = Get-MgTeamChannel -TeamId $Teamgroup.Id
    [array]$standardchannels       = $Teamchannels | Where-Object { $_.membershipType -eq "standard" }
    [array]$privatechannels        = $Teamchannels | Where-Object { $_.membershipType -eq "private" }
    [array]$outgoingsharedchannels = $Teamchannels | Where-Object { ($_.membershipType -eq "shared") -and (($_.WebUrl)  -like "*$($teamgroup.id)*") }
    [array]$incomingsharedchannels = $Teamchannels | Where-Object { ($_.membershipType -eq "shared") -and ($_.WebURL -notlike "*$($teamgroup.id)*") }

    $teamgroup | Add-Member -MemberType NoteProperty -Name "StandardChannels"       -Value $standardchannels.id.count       -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannels"        -Value $privatechannels.id.count        -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannels"         -Value $outgoingsharedchannels.id.count -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "IncomingSharedChannels" -Value $incomingsharedchannels.id.count -Force

    $privatechannelSize = 0
    foreach ($Privatechannel in $privatechannels) {
        $PrivateChannelObject = $null
        try {
            $PrivatechannelObject = Get-MgTeamChannelFileFolder -TeamId $teamgroup.id -ChannelId $Privatechannel.id
            $Privatechannelsize  += $PrivateChannelObject.size
        }
        catch {
            $Privatechannelsize  += 0
        }
    }

    $sharedchannelSize = 0
    foreach ($sharedchannel in $outgoingsharedchannels) {
        $sharedChannelObject = $null
        try {
            $SharedChannelObject = Get-MgTeamChannelFileFolder -TeamId $teamgroup.id -ChannelId $sharedChannel.id
            $Sharedchannelsize  += $SharedChannelObject.size
        }
        catch {
            $Sharedchannelsize  += 0
        }
    }

    $teamgroup | Add-Member -MemberType NoteProperty -Name "PrivateChannelsSize" -Value $privatechannelSize -Force
    $teamgroup | Add-Member -MemberType NoteProperty -Name "SharedChannelsSize"  -Value $sharedchannelSize  -Force

    $TeamDetails = $null
    try {
        [array]$TeamDetails = Get-MgGroupDrive -GroupId $teamgroup.id -ErrorAction Stop
        $teamgroup | Add-Member -MemberType NoteProperty -Name "DocumentLibraries" -Value $TeamDetails.count                          -Force
        $teamgroup | Add-Member -MemberType NoteProperty -Name "DataSize"          -Value ($TeamDetails.quota.used | Measure-Object -Sum).Sum -Force
        ##NOTE: Change for Non-English Tenants
        $teamgroup | Add-Member -MemberType NoteProperty -Name "URL"               -Value $TeamDetails[0].webUrl.Replace("/Shared%20Documents", "") -Force
    }
    catch {
        # Optionally log error
    }
}

$ProgressStatus = "Getting licenses..."
UpdateProgress
$ProgressTracker++

##Get All License SKUs
[array]$SKUs = Get-MgSubscribedSku -All

$ProgressStatus = "Getting organization details..."
UpdateProgress
$ProgressTracker++

##Get Org Details
[array]$OrgDetails = Get-MgOrganization -All

$ProgressStatus = "Getting apps..."
UpdateProgress
$ProgressTracker++

##List All Azure AD Service Principals
[array]$AADApps = Get-MgServicePrincipal -All

foreach ($user in $users) {
    $user | Add-Member -MemberType NoteProperty -Name "License SKUs"              -Value ($user.licenseAssignmentStates.skuid           -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Group License Assignments" -Value ($user.licenseAssignmentStates.assignedByGroup -join ";") -Force
    $user | Add-Member -MemberType NoteProperty -Name "Disabled Plan IDs"         -Value ($user.licenseAssignmentStates.disabledplans   -join ";") -Force
}

##Translate License SKUs and groups
foreach ($user in $users) {

    foreach ($Group in $Groups) {
        $user.'Group License Assignments' = $user.'Group License Assignments'.Replace($group.id, $group.displayName) 
    }
    foreach ($SKU in $SKUs) {
        $user.'License SKUs' = $user.'License SKUs'.Replace($SKU.skuid, $SKU.skuPartNumber)
    }
    foreach ($SKUplan in $SKUs.servicePlans) {
        $user.'Disabled Plan IDs' = $user.'Disabled Plan IDs'.Replace($SKUplan.servicePlanId, $SKUplan.servicePlanName)
    }

}

$ProgressStatus = "Getting Conditional Access policies..."
UpdateProgress
$ProgressTracker++

##Get Conditional Access Policies
[array]$ConditionalAccessPolicies = Get-MgIdentityConditionalAccessPolicy -All

##Get Directory Roles
[array]$DirectoryRoleTemplates = Get-MgDirectoryRoleTemplate

##Get Trusted Locations
[array]$NamedLocations = Get-MgIdentityConditionalAccessNamedLocation

##Tidy GUIDs to names
$ConditionalAccessPoliciesJSON = $ConditionalAccessPolicies | ConvertTo-Json -Depth 5
if ($ConditionalAccessPoliciesJSON -ne $null) {
    ##TidyUsers
    foreach ($User in $Users) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($user.id, ("$($user.displayname) - $($user.userPrincipalName)"))
    }

    ##Tidy Groups
    foreach ($Group in $Groups) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($group.id, ("$($group.displayname) - $($group.id)"))
    }

    ##Tidy Roles
    foreach ($DirectoryRoleTemplate in $DirectoryRoleTemplates) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($DirectoryRoleTemplate.Id, $DirectoryRoleTemplate.displayname)
    }

    ##Tidy Apps
    foreach ($AADApp in $AADApps) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.appid, $AADApp.displayname)
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.id,   $AADApp.displayname)
    }

    ##Tidy Locations
    foreach ($NamedLocation in $NamedLocations) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($NamedLocation.id, $NamedLocation.displayname)
    }

    $ConditionalAccessPolicies = $ConditionalAccessPoliciesJSON | ConvertFrom-Json

    $CAOutput   = @()
    $CAHeadings = @(
        "displayName",
        "createdDateTime",
        "modifiedDateTime",
        "state",
        "Conditions.users.includeusers",
        "Conditions.users.excludeusers",
        "Conditions.users.includegroups",
        "Conditions.users.excludegroups",
        "Conditions.users.includeroles",
        "Conditions.users.excluderoles",
        "Conditions.clientApplications.includeServicePrincipals",
        "Conditions.clientApplications.excludeServicePrincipals",
        "Conditions.applications.includeApplications",
        "Conditions.applications.excludeApplications",
        "Conditions.applications.includeUserActions",
        "Conditions.applications.includeAuthenticationContextClassReferences",
        "Conditions.userRiskLevels",
        "Conditions.signInRiskLevels",
        "Conditions.platforms.includePlatforms",
        "Conditions.platforms.excludePlatforms",
        "Conditions.locations.includLocations",
        "Conditions.locations.excludeLocations",
        "Conditions.clientAppTypes",
        "Conditions.devices.deviceFilter.mode",
        "Conditions.devices.deviceFilter.rule",
        "GrantControls.operator",
        "grantcontrols.builtInControls",
        "grantcontrols.customAuthenticationFactors",
        "grantcontrols.termsOfUse",
        "SessionControls.disableResilienceDefaults",
        "SessionControls.applicationEnforcedRestrictions",
        "SessionControls.persistentBrowser",
        "SessionControls.cloudAppSecurity",
        "SessionControls.signInFrequency"
    )

    foreach ($Heading in $CAHeadings) {
        $Row = New-Object psobject -Property @{
            PolicyName = $Heading
        }
    
        foreach ($CAPolicy in $ConditionalAccessPolicies) {
            $Nestingcheck = ($Heading.Split('.').Count)

            if ($Nestingcheck -eq 1) {
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value $CAPolicy.$Heading -Force
            }
            elseif ($Nestingcheck -eq 2) {
                $SplitHeading = $Heading.Split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0]).($SplitHeading[1]) -join ';') -Force
            }
            elseif ($Nestingcheck -eq 3) {
                $SplitHeading = $Heading.Split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0]).($SplitHeading[1]).($SplitHeading[2]) -join ';') -Force
            }
            elseif ($Nestingcheck -eq 4) {
                $SplitHeading = $Heading.Split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0]).($SplitHeading[1]).($SplitHeading[2]).($SplitHeading[3]) -join ';') -Force       
            }
        }

        $CAOutput += $Row
    }
}

$ProgressStatus = "Getting OneDrive report..."
UpdateProgress
$ProgressTracker++

# Get OneDrive report via direct Graph REST
try {
    $reportUri   = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')"
    $csvContent  = Invoke-RestMethod -Headers @{ Authorization = "Bearer $GraphAccessTokenPlain" } -Uri $reportUri -Method Get
    $OneDrive    = $csvContent | ConvertFrom-Csv
}
catch {
    Write-Host "Failed to get OneDrive usage report: $($_.Exception.Message)" -ForegroundColor Yellow
    $OneDrive = @()
}

$ProgressStatus = "Getting SharePoint report..."
UpdateProgress
$ProgressTracker++

# Get SharePoint report via direct Graph REST
try {
    $reportUri   = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
    $csvContent  = Invoke-RestMethod -Headers @{ Authorization = "Bearer $GraphAccessTokenPlain" } -Uri $reportUri -Method Get
    $SharePoint  = $csvContent | ConvertFrom-Csv
}
catch {
    Write-Host "Failed to get SharePoint site usage report: $($_.Exception.Message)" -ForegroundColor Yellow
    $SharePoint = @()
}

# Fallback: if usage report is unavailable/empty, enumerate sites directly
if (-not $SharePoint -or $SharePoint.Count -eq 0) {
    Write-Host "SharePoint usage report not available or empty, falling back to Get-MgSite -All..." -ForegroundColor Yellow
    $spSites = Get-MgSite -All | Where-Object { $_.WebUrl -ne $null }

    $SharePoint = foreach ($s in $spSites) {
        [pscustomobject]@{
            'Site ID'             = $s.Id
            'Site URL'            = $s.WebUrl
            'Owner Display Name'  = $null
            'Is Deleted'          = $null
            'Last Activity Date'  = $null
            'File Count'          = $null
            'Active File Count'   = $null
            'Page View Count'     = $null
            'Storage Used (Byte)' = $null
            'Root Web Template'   = $null
            'Owner Principal Name'= $null
            'TeamID'              = $null
        }
    }
}

$SharePoint | Add-Member -MemberType NoteProperty -Name "TeamID" -Value "" -Force

foreach ($Site in $Sharepoint) {
    ##NOTE: Change for Non-English Tenants
    $DriveLookup = (Get-MgSiteDrive -siteId $Site.'Site Id' -ErrorAction SilentlyContinue | Where-Object { $_.name -eq "Documents" }).weburl
    if ($DriveLookup) {
        ##NOTE: Change for Non-English Tenants
        $Site.'Site URL' = $DriveLookup.Replace('/Shared%20Documents', '')
    }
    $Site.TeamID = ($TeamGroups | Where-Object { $_.url -contains $site.'site url' }).id
}

$ProgressStatus = "Getting Mailbox Usage report..."
UpdateProgress
$ProgressTracker++

# Get Mailbox report via direct Graph REST
try {
    $reportUri          = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D30')"
    $csvContent         = Invoke-RestMethod -Headers @{ Authorization = "Bearer $GraphAccessTokenPlain" } -Uri $reportUri -Method Get
    $MailboxStatsReport = $csvContent | ConvertFrom-Csv
}
catch {
    Write-Host "Failed to get Mailbox usage report: $($_.Exception.Message)" -ForegroundColor Yellow
    $MailboxStatsReport = @()
}

$ProgressStatus = "Getting M365 Apps usage report..."
UpdateProgress

# Get M365 Apps usage report via direct Graph REST
try {
    $reportUri     = "https://graph.microsoft.com/v1.0/reports/getOffice365ServicesUserCounts(period='D30')"
    $csvContent    = Invoke-RestMethod -Headers @{ Authorization = "Bearer $GraphAccessTokenPlain" } -Uri $reportUri -Method Get
    $M365AppsUsage = $csvContent | ConvertFrom-Csv
}
catch {
    Write-Host "Failed to get Office 365 services user counts report: $($_.Exception.Message)" -ForegroundColor Yellow
    $M365AppsUsage = @()
}

##Process Group Membership
if ($IncludeGroupMembership) {
    $ProgressStatus = "Enumerating Group Membership - This may take some time..."
    UpdateProgress
    $GroupMembersObject = @()
    $i = 1
    foreach ($group in $groups) {
        $ProgressStatus = "Enumerating Group Membership - This may take some time... Processing Group $i of $($Groups.count)"
        UpdateProgress
        $i++
        $Members = Get-MgGroupMember -GroupId $group.id -All
        foreach ($member in $members) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.AdditionalProperties["displayName"]
                MemberUserPrincipalName = $member.AdditionalProperties["userPrincipalName"]
                MemberType              = "Member"
                MemberObjectType        = $member.AdditionalProperties["@odata.type"].Replace('#microsoft.graph.', '')
            }

            $GroupMembersObject += $memberEntry
        }

        $Owners = Get-MgGroupOwner -GroupId $group.id -All
        foreach ($member in $Owners) {

            $MemberEntry = [PSCustomObject]@{
                GroupID                 = $group.id
                GroupName               = $group.displayname
                MemberID                = $member.id
                MemberName              = $member.AdditionalProperties["displayName"]
                MemberUserPrincipalName = $member.AdditionalProperties["userPrincipalName"]
                MemberType              = "Owner"
                MemberObjectType        = $member.AdditionalProperties["@odata.type"].Replace('#microsoft.graph.', '')
            }

            $GroupMembersObject += $MemberEntry
        }
    }

    $ProgressTracker++
}

if ($IncludeDocumentLibraries) {
    $ProgressStatus = "Enumerating Document Libraries - This may take some time..."
    UpdateProgress
    $Sites         = Get-MgSite -All | Where-Object { $_.weburl -notlike "*sites/appcatalog*" -and $_.weburl -notlike "*sites/recordscenter*" -and $_.weburl -notlike "*sites/search*" -and $_.weburl -notlike "*sites/CompliancePolicyCenter" }
    $LibraryOutput = @()
    foreach ($site in $sites) {
        ##NOTE: Change for Non-English Tenants
        [array]$Drives = Get-MgSiteDrive -SiteId $site.id | Where-Object { $_.Name -eq "Documents" }
        foreach ($drive in $drives) {
            $LibraryObject = [PSCustomObject]@{
                LibraryID    = $Drive.id
                LibraryName  = $Drive.Name
                LibraryURL   = $Drive.WebUrl
                LibraryUsage = $Drive.quota.used
                SiteID       = $Site.id
                SiteName     = $Site.DisplayName
                SiteURL      = $Site.WebURL
            }
            $LibraryOutput += $LibraryObject
        }
    }
    $ProgressTracker++
}

if ($IncludeLists) {
    $ProgressStatus = "Enumerating Lists - This may take some time..."
    UpdateProgress
    $Sites      = Get-MgSite -All | Where-Object { $_.weburl -notlike "*sites/appcatalog*" -and $_.weburl -notlike "*sites/recordscenter*" -and $_.weburl -notlike "*sites/search*" -and $_.weburl -notlike "*sites/CompliancePolicyCenter" }
    $ListOutput = @()
    foreach ($site in $sites) {
        [array]$Lists = Get-MgSiteList -SiteId $site.id | Where-Object { $_.List.template -ne "documentLibrary" }
        foreach ($list in $lists) {
            $ListObject = [PSCustomObject]@{
                ListID   = $list.id
                ListName = $List.DisplayName
                ListURL  = $List.webUrl
                SiteID   = $Site.id
                SiteName = $Site.DisplayName
                SiteURL  = $Site.WebURL
            }
            $ListOutput += $ListObject
        }
    }
    $ProgressTracker++
}

if ($IncludePlans) {
    $ProgressStatus = "Enumerating Planner Plans - This may take some time..."
    UpdateProgress
    $unifiedGroups = $Groups | Where-Object { $_.grouptypes -Contains "unified" }
    $PlanOutput    = @()
    $PlanNumber    = 1
    foreach ($unifiedgroup in $unifiedGroups) {
        $ProgressStatus = "Enumerating Planner Plans for Group $PlanNumber of $($unifiedgroups.count) -  $($unifiedgroup.displayname)..."
        UpdateProgress
        $PlanNumber++
        [array]$Plans = Get-MgGroupPlannerPlan -GroupId $unifiedgroup.id
        foreach ($plan in $plans) {
            $PlanObject = [PSCustomObject]@{
                PlanID    = $plan.id
                PlanName  = $plan.title
                GroupID   = $unifiedgroup.id
                GroupName = $unifiedgroup.displayName
            }
            $PlanOutput += $PlanObject
        }
    }
}
$ProgressTracker++

##Tidy up Proxyaddresses
foreach ($user in $users) {
    $user | Add-member -MemberType NoteProperty -Name "Email Addresses" -Value ($user.proxyaddresses -join ';') -Force
}
foreach ($group in $groups) {
    $group | Add-member -MemberType NoteProperty -Name "Email Addresses" -Value ($group.proxyaddresses -join ';') -Force
}

###################EXCHANGE ONLINE############################

$ProgressStatus = "Connecting to Exchange Online..."
UpdateProgress
$ProgressTracker++

try {
    Connect-ExchangeOnline -Certificate $Certificate -AppID $clientid -Organization ($orgdetails.verifieddomains | Where-Object { $_.isinitial -eq "true" }).name -ShowBanner:$false
}
catch {
    Write-Host "Error connecting to Exchange Online...Exiting..." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Pause
    Exit
}

$ProgressStatus = "Getting shared and room mailboxes..."
UpdateProgress
$ProgressTracker++

[array]$RoomMailboxes      = Get-EXOMailbox -RecipientTypeDetails RoomMailbox      -ResultSize unlimited
[array]$EquipmentMailboxes = Get-EXOMailbox -RecipientTypeDetails EquipmentMailbox -ResultSize unlimited
[array]$SharedMailboxes    = Get-EXOMailbox -RecipientTypeDetails SharedMailbox    -ResultSize Unlimited

$ProgressStatus = "Getting room mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1
foreach ($room in $RoomMailboxes) {
    $ProgressStatus = "Getting room mailbox statistics $i of $($RoomMailboxes.count)..."
    $i++
    UpdateProgress

    $RoomStats = Get-EXOMailboxStatistics $room.primarysmtpaddress
    $room | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $RoomStats.TotalItemSize -Force
    $room | Add-Member -MemberType NoteProperty -Name ItemCount   -Value $RoomStats.ItemCount     -Force

    $room.EmailAddresses = $room.EmailAddresses -join ';'
}

$ProgressStatus = "Getting Equipment mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1
foreach ($equipment in $EquipmentMailboxes) {
    $ProgressStatus = "Getting Equipment mailbox statistics $i of $($EquipmentMailboxes.count)..."
    $i++
    UpdateProgress

    $EquipmentStats = Get-EXOMailboxStatistics $equipment.primarysmtpaddress
    $equipment | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $EquipmentStats.TotalItemSize -Force
    $equipment | Add-Member -MemberType NoteProperty -Name ItemCount   -Value $EquipmentStats.ItemCount     -Force

    $equipment.EmailAddresses = $equipment.EmailAddresses -join ';'
}

$ProgressStatus = "Getting shared mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 1
foreach ($SharedMailbox in $SharedMailboxes) {
    $ProgressStatus = "Getting shared mailbox statistics $i of $($SharedMailboxes.count)..."
    $i++
    UpdateProgress

    $SharedStats = Get-EXOMailboxStatistics $SharedMailbox.primarysmtpaddress
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name MailboxSize -Value $SharedStats.TotalItemSize -Force
    $SharedMailbox | Add-Member -MemberType NoteProperty -Name ItemCount   -Value $SharedStats.ItemCount     -Force
    
    $SharedMailbox.EmailAddresses = $SharedMailbox.EmailAddresses -join ';'
}

$ProgressStatus = "Getting user mailbox statistics..."
UpdateProgress
$ProgressTracker++

##Collect Mailbox statistics
$MailboxStats = @()
foreach ($user in ($users | Where-Object { ($_.mail -ne $null ) -and ($_.userType -eq "Member") })) {
    $stats = $MailboxStatsReport | Where-Object { $_.'User Principal Name' -eq $user.userprincipalname }
    if ($stats) {
        $stats | Add-Member -MemberType NoteProperty -Name ObjectID           -Value $user.id   -Force
        $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $user.mail -Force
        $MailboxStats += $stats
    }
}

$ProgressStatus = "Getting archive mailbox statistics..."
UpdateProgress
$ProgressTracker++

$i = 0

##Collect Archive Statistics
$ArchiveStats      = @()
[array]$ArchiveMailboxes = Get-EXOMailbox -Archive -ResultSize unlimited
foreach ($archive in $ArchiveMailboxes) {
    $ProgressStatus = "Getting archive mailbox statistics $i of $($ArchiveMailboxes.count)..."
    $i++
    UpdateProgress
    $stats = Get-EXOMailboxStatistics $archive.PrimarySmtpAddress -Archive
    $stats | Add-Member -MemberType NoteProperty -Name ObjectID           -Value $archive.ExternalDirectoryObjectId -Force
    $stats | Add-Member -MemberType NoteProperty -Name Primarysmtpaddress -Value $archive.primarysmtpaddress       -Force
    $ArchiveStats += $stats
}

$ProgressStatus = "Getting mail contacts..."
UpdateProgress
$ProgressTracker++

##Collect Mail Contacts
$MailContacts = Get-MailContact -ResultSize unlimited | Select-Object displayname, alias, externalemailaddress, emailaddresses, HiddenFromAddressListsEnabled
foreach ($mailcontact in $MailContacts) {
    $mailcontact.emailaddresses = $mailcontact.emailaddresses -join ';'
}

$ProgressStatus = "Getting transport rules..."
UpdateProgress
$ProgressTracker++

##Collect transport rules
[array]$Rules = Get-TransportRule -ResultSize unlimited | Select-Object name, state, mode, priority, description, comments
$RulesOutput  = @()
foreach ($Rule in $Rules) {
    $RulesOutput += $Rule
}

#######Optional Items - EXO#######

##Process Mailbox Permissions
if ($IncludeMailboxPermissions) {
    $ProgressStatus = "Fetching Mailbox Permissions - This may take some time..."
    UpdateProgress
    $PermissionOutput   = @()
    $MailboxList        = Get-EXOMailbox -ResultSize unlimited
    $PermissionProgress = 1
    foreach ($mailbox in $MailboxList) {
        $ProgressStatus = "Fetching Mailbox Permissions for mailbox $PermissionProgress of $($Mailboxlist.count) - This may take some time..."
        UpdateProgress

        [array]$Permissions = Get-EXOMailboxPermission -UserPrincipalName $mailbox.UserPrincipalName | Where-Object { $_.User -ne "NT AUTHORITY\SELF" }

        foreach ($permission in $Permissions) {
            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.user
            }
            
            $PermissionOutput += $PermissionObject
        }

        [array]$RecipientPermissions = Get-EXORecipientPermission $mailbox.UserPrincipalName | Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" }

        foreach ($permission in $RecipientPermissions) {
            $PermissionObject = [PSCustomObject]@{
                ExternalDirectoryObjectId = $mailbox.ExternalDirectoryObjectId
                UserPrincipalName         = $Mailbox.UserPrincipalName
                Displayname               = $mailbox.DisplayName
                PrimarySmtpAddress        = $mailbox.PrimarySmtpAddress
                AccessRight               = $permission.accessRights -join ';'
                GrantedTo                 = $Permission.trustee
            }
            
            $PermissionOutput += $PermissionObject
        }

        $PermissionProgress++
    }
    $ProgressTracker++
}

#######Report Export#######

$ProgressStatus = "Getting mail connectors..."
UpdateProgress
$ProgressTracker++

##Collect Mailflow Connectors

$InboundConnectors = Get-InboundConnector | Select-Object enabled, name, connectortype, connectorsource, SenderIPAddresses, SenderDomains, RequireTLS, RestrictDomainsToIPAddresses, RestrictDomainsToCertificate, CloudServicesMailEnabled, TreatMessagesAsInternal, TlsSenderCertificateName, EFTestMode, Comment 
foreach ($inboundconnector in $InboundConnectors) {
    $inboundconnector.senderipaddresses = $inboundconnector.senderipaddresses -join ';'
    $inboundconnector.senderdomains     = $inboundconnector.senderdomains     -join ';'
}
$OutboundConnectors = Get-OutboundConnector -IncludeTestModeConnectors:$true | Select-Object enabled, name, connectortype, connectorsource, TLSSettings, RecipientDomains, UseMXRecord, SmartHosts, Comment
foreach ($OutboundConnector in $OutboundConnectors) {
    $OutboundConnector.RecipientDomains = $OutboundConnector.RecipientDomains -join ';'
    $OutboundConnector.SmartHosts       = $OutboundConnector.SmartHosts       -join ';'
}

$ProgressStatus = "Getting MX records..."
UpdateProgress
$ProgressTracker++

##MX Record Check
$MXRecordsObject = @()
foreach ($domain in $orgdetails.verifieddomains) {
    try {
        [array]$MXRecords = Resolve-DnsName -Name $domain.name -Type mx -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "Error obtaining MX Record for $($domain.name)"
    }
    foreach ($MXRecord in $MXRecords) {
        $MXRecordsObject += $MXRecord
    }
}

$ProgressStatus = "Updating references..."
UpdateProgress
$ProgressTracker++

##Update users tab with Values
$users | Add-Member -MemberType NoteProperty -Name MailboxSizeGB     -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name MailboxItemCount  -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveSizeGB    -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name OneDriveFileCount -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveSizeGB     -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name Mailboxtype       -Value "" -Force
$users | Add-Member -MemberType NoteProperty -Name ArchiveItemCount  -Value "" -Force

foreach ($user in ($users | Where-Object { $_.usertype -ne "Guest" })) {
    ##Set Mailbox Type
    if     ($roommailboxes.ExternalDirectoryObjectId      -contains $user.id) { $user.Mailboxtype = "Room" }
    elseif ($EquipmentMailboxes.ExternalDirectoryObjectId -contains $user.id) { $user.Mailboxtype = "Equipment" }
    elseif ($sharedmailboxes.ExternalDirectoryObjectId    -contains $user.id) { $user.Mailboxtype = "Shared" }
    else                                                                      { $user.Mailboxtype = "User" }

    ##Set Mailbox Size and count (User mailbox stats)
    $thisMbx = $MailboxStats | Where-Object { $_.objectID -eq $user.id }
    if ($thisMbx) {
        $user.MailboxSizeGB    = [math]::Round((($thisMbx.'Storage Used (Byte)' / 1024 / 1024 / 1024)), 2)
        $user.MailboxItemCount = $thisMbx.'item count'
    }

    ##Set Shared Mailbox size and count
    $thisShared = $SharedMailboxes | Where-Object { $_.ExternalDirectoryObjectId -eq $user.id }
    if ($thisShared -and $thisShared.mailboxsize) {
        $user.MailboxSizeGB    = [math]::Round((($thisShared.mailboxsize.value.ToString().Replace(',', '').Replace(' ', '').Split('b')[0].Split('(')[1] / 1024 / 1024 / 1024)), 2)
        $user.MailboxItemCount = $thisShared.ItemCount
    }

    ##Set Equipment Mailbox size and count
    $thisEquip = $EquipmentMailboxes | Where-Object { $_.ExternalDirectoryObjectId -eq $user.id }
    if ($thisEquip -and $thisEquip.mailboxsize) {
        $user.MailboxSizeGB    = [math]::Round((($thisEquip.mailboxsize.value.ToString().Replace(',', '').Replace(' ', '').Split('b')[0].Split('(')[1] / 1024 / 1024 / 1024)), 2)
        $user.MailboxItemCount = $thisEquip.ItemCount
    }

    ##Set Room Mailbox size and count
    $thisRoom = $roommailboxes | Where-Object { $_.ExternalDirectoryObjectId -eq $user.id }
    if ($thisRoom -and $thisRoom.mailboxsize) {
        $user.MailboxSizeGB    = [math]::Round((($thisRoom.mailboxsize.value.ToString().Replace(',', '').Replace(' ', '').Split('b')[0].Split('(')[1] / 1024 / 1024 / 1024)), 2)
        $user.MailboxItemCount = $thisRoom.ItemCount
    }

    ##Set archive size and count
    $thisArchive = $ArchiveStats | Where-Object { $_.objectID -eq $user.id }
    if ($thisArchive) {
        $user.ArchiveSizeGB    = [math]::Round((($thisArchive.totalitemsize.value.ToString().Replace(',', '').Replace(' ', '').Split('b')[0].Split('(')[1] / 1024 / 1024 / 1024)), 2)
        $user.ArchiveItemCount = $thisArchive.ItemCount
    }

    ##Set OneDrive Size and count
    $thisOD = $OneDrive | Where-Object { $_.'Owner Principal Name' -eq $user.UserPrincipalName }
    if ($thisOD -and $thisOD.'Storage Used (Byte)') {
        $user.OneDriveSizeGB    = [math]::Round((($thisOD.'Storage Used (Byte)' / 1024 / 1024 / 1024)), 2)
        $user.OneDriveFileCount = $thisOD.'file count'
    }
}

$ProgressStatus = "Exporting report..."
UpdateProgress
$ProgressTracker++

try {
    if ($TemplatePresent) {
        ##Add cover sheet
        Copy-ExcelWorksheet -SourceObject TenantAssessment-Template.xlsx -SourceWorksheet "High-Level" -DestinationWorkbook "$FilePath\$Filename" -DestinationWorksheet "High-Level"
    }

    $users      | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $SharePoint | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $TeamGroups | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force
    $Groups     | Add-Member -MemberType NoteProperty -Name "Migrate" -Value "TRUE" -Force

    ##Export Data File##
    ##Export User Accounts tab
    $users | Where-Object { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "User") } |
        Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, OneDriveSizeGB, OneDriveFileCount, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "User Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 

    ##Export Shared Mailboxes tab
    $users | Where-Object { ($_.usertype -ne "Guest") -and ($_.mailboxtype -eq "shared") } |
        Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Shared Mailboxes" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 

    ##Export Resource Accounts tab
    $users | Where-Object { ($_.usertype -ne "Guest") -and (($_.mailboxtype -eq "Room") -or ($_.mailboxtype -eq "Equipment")) } |
        Select-Object Migrate, id, accountenabled, userPrincipalName, mail, targetobjectID, targetUPN, TargetMail, displayName, MailboxItemCount, MailboxSizeGB, MailboxType, ArchiveSizeGB, ArchiveItemCount, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usagelocation, usertype, onPremisesSyncEnabled  |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Resource Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 

    ##Export SharePoint Tab
    $SharePoint | Where-Object { ($_.teamid -eq $null) -and ($_.'Root Web Template' -ne "Team Channel") } |
        Select-Object Migrate, 'Site ID', 'Site URL', 'Owner Display Name', 'Is Deleted', 'Last Activity Date', 'File Count', 'Active File Count', 'Page View Count', 'Storage Used (Byte)', 'Root Web Template', 'Owner Principal Name' |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "SharePoint Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Teams Tab
    $TeamGroups |
        Select-Object Migrate, id, displayname, standardchannels, privatechannels, SharedChannels, Datasize, PrivateChannelsSize, SharedChannelsSize, IncomingSharedChannels, mail, URL, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Teams"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Unified Groups tab
    $Groups | Where-Object { ($_.grouptypes -Contains "unified") -and ($_.resourceProvisioningOptions -notcontains "Team") } |
        Select-Object Migrate, id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility, membershipRule |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Unified Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Standard Groups tab
    $Groups | Where-Object { $_.grouptypes -notContains "unified" } |
        Select-Object Migrate, id, displayname, mail, description, createdDateTime, mailEnabled, securityenabled, mailNickname, 'Email Addresses', visibility, membershipRule |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Standard Groups"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Guest Accounts tab
    $users | Where-Object { $_.usertype -eq "Guest" } |
        Select-Object id, accountenabled, userPrincipalName, mail, displayName, givenName, surname, 'Email Addresses', 'License SKUs', 'Group License Assignments', 'Disabled Plan IDs', usertype |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Guest Accounts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow 

    ##Export AAD Apps Tab
    $AADApps | Where-Object { $_.publishername -notlike "Microsoft*" } |
        Select-Object createddatetime, displayname, publisherName, signinaudience |
        Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "AAD Apps" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Conditional Access Tab
    $CAOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Conditional Access" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export M365 Apps Usage
    $M365AppsUsage | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "M365 Apps Usage" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Mail Contacts tab
    $MailContacts | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MailContacts" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export MX Records tab
    $MXRecordsObject | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "MX Records"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Verified Domains tab
    $orgdetails.verifieddomains | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Verified Domains"  -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Transport Rules tab
    $RulesOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Transport Rules" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Receive Connectors Tab
    $InboundConnectors | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Receive Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export Send Connectors Tab
    $OutboundConnectors | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Send Connectors" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    ##Export OneDrive Tab
    $OneDrive | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "OneDrive Sites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

    if ($IncludeMailboxPermissions) {
        ##Export Mailbox Permissions Tab
        $PermissionOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Mailbox Permissions" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    if ($IncludeGroupMembership) {
        ##Export Group Membership Tab
        $GroupMembersObject | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Group Membership" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    if ($IncludeDocumentLibraries) {
        ##Export Document Libraries Tab
        $LibraryOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Document Libraries" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    if ($IncludeLists) {
        ##Export Lists Tab
        $ListOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Lists" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    if ($IncludePlans) {
        ##Export Planner Plans Tab
        $PlanOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Planner Plans" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
}
catch {
    Write-Host "Error exporting report, check permissions and make sure the file is not open! $_"
    Pause
}

$ProgressStatus = "Finalizing..."
UpdateProgress
$ProgressTracker++

# === Runtime timer: record end time and print summary ===
$ScriptEndTime = Get-Date
$RunDuration   = $ScriptEndTime - $ScriptStartTime
$RunMinutes    = [math]::Round($RunDuration.TotalMinutes, 2)

Write-Host ""
Write-Host "Tenant Assessment completed." -ForegroundColor Green
Write-Host ("Runtime: {0:hh\:mm\:ss} (~{1} minutes)" -f $RunDuration, $RunMinutes) -ForegroundColor Green
Write-Host "Start: $ScriptStartTime" -ForegroundColor DarkGray
Write-Host "End  : $ScriptEndTime"   -ForegroundColor DarkGray
Write-Host "Output file: $FilePath\$Filename" -ForegroundColor Cyan

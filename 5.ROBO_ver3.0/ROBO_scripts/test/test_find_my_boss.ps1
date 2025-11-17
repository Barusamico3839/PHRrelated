#
# test_find_my_boss.ps1
# Fetches the specified user's profile and their manager from Microsoft Graph.
# Default sign-in uses the interactive browser flow (quickest path when loopback works).
# Use -UseDeviceAuth to switch to the device-code flow when localhost redirect is blocked.
#
# Requirements:
# - PowerShell 7+
# - Microsoft.Graph PowerShell SDK (installed automatically for the current user if missing)
# - Delegated permissions User.Read.All and Directory.Read.All (admin consent typically required)
#
# Example:
#   ./test_find_my_boss.ps1 -User "okada.kazuhito@jp.panasonic.com"
#   ./test_find_my_boss.ps1 -User "okada.kazuhito@jp.panasonic.com" -UseDeviceAuth
#   ./test_find_my_boss.ps1 -User "okada.kazuhito@jp.panasonic.com" -Raw
#
param(
    [Parameter(Mandatory = $false)]
    [string]$User = "okada.kazuhito@jp.panasonic.com",

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceAuth,

    [Parameter(Mandatory = $false)]
    [switch]$Raw,

    [Parameter(Mandatory = $false)]
    [int]$RequestTimeoutSeconds = 3
)

$ErrorActionPreference = 'Stop'
$script:UseDeviceAuth = $UseDeviceAuth
$script:RawOutput = $Raw
$script:RequestTimeoutSeconds = [Math]::Max($RequestTimeoutSeconds, 1)

Write-Host ("[STEP] Starting manager lookup for {0}" -f $User) -ForegroundColor Green

$script:UserSelectProperties = @(
    'id','displayName','mail','userPrincipalName','jobTitle','department','companyName',
    'businessPhones','mobilePhone','officeLocation','preferredLanguage','givenName','surname',
    'mailNickname','userType','accountEnabled','otherMails','imAddresses',
    'employeeId','employeeType','employeeHireDate','employeeOrgData','createdDateTime',
    'onPremisesSamAccountName','onPremisesUserPrincipalName','onPremisesDistinguishedName',
    'onPremisesDomainName','onPremisesImmutableId','country','city','state','postalCode','streetAddress',
    'usageLocation','preferredName','displayNamePronunciation'
)
$script:SelectQuery = ($script:UserSelectProperties -join ',')

function Ensure-GraphModule {
    Write-Host "[STEP] Ensuring Microsoft Graph module is available..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Host 'Installing Microsoft.Graph module for current user...' -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module Microsoft.Graph -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue | Out-Null
    $timeoutCmd = Get-Command -Name Set-MgGraphRequestTimeout -ErrorAction SilentlyContinue
    if ($timeoutCmd) {
        try {
            Set-MgGraphRequestTimeout -Milliseconds ($script:RequestTimeoutSeconds * 1000)
            Write-Host ("[STEP] Graph request timeout set to {0} seconds." -f $script:RequestTimeoutSeconds) -ForegroundColor DarkCyan
        } catch {
            Write-Warning "Failed to set Graph request timeout: $($_.Exception.Message)"
        }
    }
}

function Connect-GraphIfNeeded {
    $requiredScopes = @('User.Read.All','Directory.Read.All')
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
    } catch {
        $ctx = $null
    }

    $needConnect = $true
    if ($ctx) {
        $currentScopes = @($ctx.Scopes)
        $missingScopes = $requiredScopes | Where-Object { $currentScopes -notcontains $_ }
        if (-not $missingScopes) {
            $needConnect = $false
        } else {
            Write-Host ("[STEP] Existing Graph context missing scopes: {0}" -f ($missingScopes -join ', ')) -ForegroundColor Yellow
        }
    } else {
        Write-Host "[STEP] No active Graph context detected; connection required." -ForegroundColor Yellow
    }

    if ($needConnect) {
        Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Yellow
        try {
            if ($script:UseDeviceAuth) {
                Connect-MgGraph -Scopes $requiredScopes -UseDeviceAuthentication | Out-Null
            } else {
                Connect-MgGraph -Scopes $requiredScopes | Out-Null
            }
            Write-Host "[STEP] Graph authentication completed." -ForegroundColor DarkGreen
        } catch [System.Management.Automation.PipelineStoppedException] {
            throw 'Graph sign-in was interrupted (PipelineStoppedException). Re-run and complete the sign-in prompt, or retry with -UseDeviceAuth.'
        }
        $selectCmd = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
        if ($selectCmd) {
            Select-MgProfile -Name beta -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

function Get-UserManagerObject {
    param(
        [Parameter(Mandatory = $true)][string]$UserId
    )

    try {
        $managerCmd = Get-Command -Name Get-MgUserManagerUser -ErrorAction SilentlyContinue
        if ($managerCmd) {
            try {
                Write-Host "[STEP] Fetching manager via Get-MgUserManagerUser..." -ForegroundColor Cyan
                return Get-MgUserManagerUser -UserId $UserId -ErrorAction Stop
            } catch {
                # fall through to generic handler below
                Write-Host "[STEP] Get-MgUserManagerUser failed; falling back to alternate methods..." -ForegroundColor Yellow
            }
        }

        Write-Host "[STEP] Fetching manager via Get-MgUserManager..." -ForegroundColor Cyan
        $mgr = Get-MgUserManager -UserId $UserId -ErrorAction Stop
        if ($mgr -and ($mgr.PSObject.Properties.Name -contains 'AdditionalProperties') -and $mgr.AdditionalProperties) {
            $data = [ordered]@{}
            foreach ($entry in $mgr.AdditionalProperties.GetEnumerator()) {
                $data[$entry.Key] = $entry.Value
            }
            if ($mgr.PSObject.Properties.Name -contains 'Id' -and -not $data.Contains('id')) {
                $data['id'] = $mgr.Id
            }
            return [pscustomobject]$data
        }
        return $mgr
    } catch {
        $message = $_.Exception.Message
        if ($message -match 'Insufficient privileges' -or $message -match 'Authorization_RequestDenied') {
            throw 'Microsoft Graph permissions are insufficient. Request admin consent for User.Read.All and Directory.Read.All.'
        }
        if ($message -match 'Request_ResourceNotFound' -or $message -match 'ResourceNotFound' -or $message -match 'NotFound') {
            throw "User '$UserId' or their manager could not be found."
        }
        Write-Host '[STEP] Falling back to Invoke-MgGraphRequest for manager lookup...' -ForegroundColor Yellow
        $uri = "/users/{0}/manager?$select=$($script:SelectQuery)" -f $UserId
        return Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop -RetryCount 1
    }
}

function Get-UserDetail {
    param(
        [Parameter(Mandatory = $true)][string]$UserId
    )

    try {
        Write-Host "[STEP] Fetching user details via Get-MgUser..." -ForegroundColor Cyan
        return Get-MgUser -UserId $UserId -Property $script:UserSelectProperties -ErrorAction Stop
    } catch {
        $message = $_.Exception.Message
        if ($message -match 'Request_ResourceNotFound' -or $message -match 'ResourceNotFound' -or $message -match 'NotFound') {
            throw "User '$UserId' could not be found."
        }
        throw $_
    }
}

function Get-ManagerDetail {
    param(
        [Parameter(Mandatory = $true)][object]$ManagerObject
    )

    $managerId = $ManagerObject.Id
    if (-not $managerId) {
        return $null
    }

    try {
        Write-Host "[STEP] Fetching manager detail via Get-MgUser..." -ForegroundColor Cyan
        return Get-MgUser -UserId $managerId -Property $script:UserSelectProperties -ErrorAction Stop
    } catch {
        Write-Host "[STEP] Get-MgUser for manager failed; switching to Invoke-MgGraphRequest..." -ForegroundColor Yellow
        Write-Verbose 'Get-MgUser failed; retrying with Invoke-MgGraphRequest.'
    }

    try {
        $uri = "/users/{0}?$select=$($script:SelectQuery)" -f $managerId
        Write-Host "[STEP] Fetching manager detail via Invoke-MgGraphRequest..." -ForegroundColor Yellow
        return Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop -RetryCount 1
    } catch {
        Write-Warning "Unable to fetch extended manager detail: $($_.Exception.Message)"
        return $null
    }
}

function Format-PersonOutput {
    param(
        [Parameter(Mandatory = $true)][object]$Primary,
        [Parameter(Mandatory = $false)][object]$Secondary
    )

    $ordered = [ordered]@{}
    $sources = @($Primary, $Secondary) | Where-Object { $_ }
    $fields = @(
        'DisplayName','Mail','UserPrincipalName','Id','JobTitle','Department',
        'CompanyName','OfficeLocation','PreferredLanguage','GivenName','Surname',
        'MailNickname','UserType','AccountEnabled','BusinessPhones','MobilePhone',
        'OtherMails','ImAddresses','EmployeeId','EmployeeType','EmployeeHireDate','EmployeeOrgData',
        'Country','City','State','PostalCode','StreetAddress','UsageLocation',
        'OnPremisesSamAccountName','OnPremisesUserPrincipalName','OnPremisesDistinguishedName',
        'OnPremisesDomainName','OnPremisesImmutableId'
    )

    foreach ($field in $fields) {
        foreach ($src in $sources) {
            if ($src.PSObject.Properties.Name -contains $field) {
                $value = $src.$field
                if ($null -ne $value) {
                    if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                        $value = ($value -join ', ')
                    }
                    $ordered[$field] = $value
                } else {
                    $ordered[$field] = $null
                }
                break
            }
        }
    }

    return [pscustomobject]$ordered
}

function Get-UserExtendedData {
    param(
        [Parameter(Mandatory = $true)][string]$UserId
    )

    $extended = [ordered]@{}

    try {
        $extended['LicenseDetails'] = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
    } catch {
        $extended['LicenseDetails'] = @()
    }

    try {
        $extended['AppRoleAssignments'] = Get-MgUserAppRoleAssignment -UserId $UserId -ErrorAction Stop
    } catch {
        $extended['AppRoleAssignments'] = @()
    }

    try {
        $extended['MemberOf'] = Get-MgUserMemberOf -UserId $UserId -ConsistencyLevel eventual -Top 20 -ErrorAction Stop
    } catch {
        $extended['MemberOf'] = @()
    }

    try {
        $extended['AuthenticationMethods'] = Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop
    } catch {
        $extended['AuthenticationMethods'] = @()
    }

    return [pscustomobject]$extended
}

function Write-ExtendedData {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][pscustomobject]$Extended
    )

    Write-Host "" -ForegroundColor Gray
    Write-Host ("[{0}] License Details:" -f $Label) -ForegroundColor DarkCyan
    if ($Extended.LicenseDetails -and $Extended.LicenseDetails.Count) {
        $Extended.LicenseDetails | Select-Object SkuId, SkuPartNumber, AssignedPlans | Format-Table -AutoSize
    } else {
        Write-Host "  (none)"
    }

    Write-Host ("[{0}] App Role Assignments:" -f $Label) -ForegroundColor DarkCyan
    if ($Extended.AppRoleAssignments -and $Extended.AppRoleAssignments.Count) {
        $Extended.AppRoleAssignments | Select-Object ResourceDisplayName, ResourceId, AppRoleId | Format-Table -AutoSize
    } else {
        Write-Host "  (none)"
    }

    Write-Host ("[{0}] Member Of (top 20):" -f $Label) -ForegroundColor DarkCyan
    if ($Extended.MemberOf -and $Extended.MemberOf.Count) {
        $Extended.MemberOf | Select-Object Id, DisplayName, '@odata.type' | Format-Table -AutoSize
    } else {
        Write-Host "  (none)"
    }

    if ($Extended.AuthenticationMethods -and $Extended.AuthenticationMethods.Count) {
        Write-Host ("[{0}] Authentication Methods:" -f $Label) -ForegroundColor DarkCyan
        $Extended.AuthenticationMethods | Select-Object Id, '@odata.type' | Format-Table -AutoSize
    }
}

try {
    Write-Host "[STEP] Ensuring Microsoft Graph prerequisites..." -ForegroundColor Green
    Ensure-GraphModule
    Connect-GraphIfNeeded
    $userDetail = Get-UserDetail -UserId $User
    if (-not $userDetail) {
        throw 'Failed to retrieve user information.'
    }

    $managerObj = Get-UserManagerObject -UserId $User
    if (-not $managerObj) {
        throw 'Failed to retrieve manager information.'
    }

    $managerDetail = Get-ManagerDetail -ManagerObject $managerObj

    if ($script:RawOutput) {
        $payload = [ordered]@{
            User    = $userDetail
            Manager = if ($managerDetail) { $managerDetail } else { $managerObj }
        }
        $payload | ConvertTo-Json -Depth 10
        exit 0
    }

    $userOut = Format-PersonOutput -Primary $userDetail
    $managerOut = Format-PersonOutput -Primary $managerObj -Secondary $managerDetail

    Write-Host ('User: {0}' -f $User) -ForegroundColor Cyan
    $userOut | Format-List

    Write-Host ("Manager of {0}:" -f $User) -ForegroundColor Cyan
    $managerOut | Format-List

    $userExtended = Get-UserExtendedData -UserId $User
    $managerExtended = $null
    if ($managerDetail -and $managerDetail.Id) {
        $managerExtended = Get-UserExtendedData -UserId $managerDetail.Id
    } elseif ($managerObj.Id) {
        try {
            $managerExtended = Get-UserExtendedData -UserId $managerObj.Id
        } catch {
            Write-Warning "Unable to gather extended data for manager fallback object."
        }
    }

    Write-ExtendedData -Label 'User' -Extended $userExtended
    if ($managerExtended) {
        Write-ExtendedData -Label 'Manager' -Extended $managerExtended
    } else {
        Write-Host "[INFO] Manager extended data unavailable." -ForegroundColor Yellow
    }

    Write-Host "[STEP] Manager lookup completed successfully." -ForegroundColor Green
} catch {
    Write-Error $_
    exit 1
}



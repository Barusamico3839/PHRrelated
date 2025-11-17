param(
    [Parameter(Mandatory = $true)]
    [string]$UserEmail,

    [Parameter(Mandatory = $false)]
    [string[]]$Scopes = @('User.Read.All','Directory.Read.All'),

    [Parameter(Mandatory = $false)]
    [switch]$SkipModuleInstall,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeExtendedData,

    [Parameter(Mandatory = $false)]
    [int]$RequestTimeoutSeconds = 3
)

if ($Scopes.Count -eq 1 -and $Scopes[0] -match ',') {
    $Scopes = $Scopes[0].Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
}

$ErrorActionPreference = 'Stop'
$includeExtended = [bool]$IncludeExtendedData
$RequestTimeoutSeconds = [Math]::Max($RequestTimeoutSeconds, 1)

$script:ScopeCandidates = $null

$UserSelectProperties = @(
    'id','displayName','mail','userPrincipalName','jobTitle','department','companyName',
    'businessPhones','mobilePhone','officeLocation','preferredLanguage','givenName','surname',
    'mailNickname','userType','accountEnabled','otherMails','imAddresses',
    'employeeId','employeeType','employeeHireDate','employeeOrgData','createdDateTime',
    'onPremisesSamAccountName','onPremisesUserPrincipalName','onPremisesDistinguishedName',
    'onPremisesDomainName','onPremisesImmutableId','country','city','state','postalCode','streetAddress',
    'usageLocation','preferredName','displayNamePronunciation'
)

function Write-Step {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::Cyan
    )
    Write-Host "[STEP] $Message" -ForegroundColor $Color
}

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor DarkGray
}

function Add-ScopeCandidate {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [System.Collections.Generic.HashSet[string]]$Seen,
        [string[]]$Scopes
    )

    if (-not $Scopes) { return }
    $normalized = @()
    foreach ($scope in $Scopes) {
        if ([string]::IsNullOrWhiteSpace($scope)) { continue }
        $normalized += $scope.Trim()
    }
    if ($normalized.Count -eq 0) { return }
    $key = (($normalized | ForEach-Object { $_.ToLowerInvariant() }) | Sort-Object -Unique) -join '|'
    if (-not $key) { return }
    if ($Seen.Add($key)) {
        $Target.Add($normalized) | Out-Null
    }
}

function Get-GraphScopeCandidates {
    param([string[]]$PrimaryScopes)

    $candidates = New-Object System.Collections.Generic.List[object]
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes $PrimaryScopes
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes @('User.Read.All','Directory.Read.All')
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes @('Directory.Read.All')
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes @('User.Read.All')
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes @('User.ReadBasic.All')
    Add-ScopeCandidate -Target $candidates -Seen $seen -Scopes @('User.Read')
    if ($candidates.Count -eq 0) {
        $candidates.Add(@('User.Read')) | Out-Null
    }
    return ,$candidates
}

function Test-GraphConsentError {
    param([string]$Message)

    if ([string]::IsNullOrWhiteSpace($Message)) {
        return $false
    }
    $patterns = @(
        'AADSTS65001',
        'consent',
        'Need admin approval',
        'Authorization_RequestDenied',
        'Insufficient privileges',
        'does not have access'
    )
    foreach ($pattern in $patterns) {
        if ($Message -match $pattern) {
            return $true
        }
    }
    return $false
}

function Connect-GraphWithFallback {
    param([System.Collections.Generic.List[object]]$ScopeSets)

    if (-not $ScopeSets -or $ScopeSets.Count -eq 0) {
        throw 'No Graph scope sets specified.'
    }

    $lastError = $null
    foreach ($scopeSet in $ScopeSets) {
        $scopes = @($scopeSet)
        $label = if ($scopes.Count -gt 0) { $scopes -join ', ' } else { '(default)' }
        Write-Info ("Attempting Graph connection with scopes: {0}" -f $label)
        try {
            if ($scopes.Count -gt 0) {
                Connect-MgGraph -Scopes $scopes | Out-Null
            } else {
                Connect-MgGraph | Out-Null
            }
            Write-Info ("Graph connection established with scopes: {0}" -f $label)
            return $scopes
        } catch {
            $lastError = $_
            $message = $_.Exception.Message
            if (Test-GraphConsentError -Message $message) {
                Write-Step ("Consent missing for scopes {0}: {1}" -f $label, $message) Yellow
                continue
            }
            throw
        }
    }

    if ($lastError) {
        throw $lastError
    }
    throw 'Graph connection failed for all scope sets.'
}

function Test-ContextSatisfiesScopes {
    param(
        [string[]]$ContextScopes,
        [string[]]$RequiredScopes
    )

    if (-not $RequiredScopes -or $RequiredScopes.Count -eq 0) {
        return $true
    }
    if (-not $ContextScopes -or $ContextScopes.Count -eq 0) {
        return $false
    }
    foreach ($scope in $RequiredScopes) {
        if ($ContextScopes -notcontains $scope) {
            return $false
        }
    }
    return $true
}

function Resolve-RpaBookPath {
    return [System.IO.Path]::Combine(
        $env:USERPROFILE,
        "Desktop",
        "【全社標準】弔事対応フォルダ",
        "2. RPAブック",
        "RPAブック.xlsx"
    )
}

function Ensure-GraphModule {
    Write-Step "Ensuring Microsoft Graph module is available..."
    $moduleName = 'Microsoft.Graph.Authentication'
    $loaded = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
    if (-not $loaded) {
        $available = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
        if (-not $available) {
            if ($SkipModuleInstall) {
                throw "Microsoft.Graph.Authentication module is not installed. Run without -SkipModuleInstall or install manually."
            }
            Write-Step "Installing Microsoft.Graph.Authentication for current user..." Yellow
            Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue | Out-Null
    $timeoutCmd = Get-Command -Name Set-MgGraphRequestTimeout -ErrorAction SilentlyContinue
    if ($timeoutCmd) {
        try {
            Set-MgGraphRequestTimeout -Milliseconds ($RequestTimeoutSeconds * 1000)
            Write-Info ("Graph timeout set to {0} seconds." -f $RequestTimeoutSeconds)
        } catch {
            Write-Info ("Unable to set Graph timeout: {0}" -f $_.Exception.Message)
        }
    }
}

function Connect-GraphIfNeeded {
    param([string[]]$DesiredScopes)

    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
    } catch {
        $ctx = $null
    }

    $scopeCandidates = $script:ScopeCandidates
    if (-not $scopeCandidates) {
        $scopeCandidates = Get-GraphScopeCandidates -PrimaryScopes $DesiredScopes
        $script:ScopeCandidates = $scopeCandidates
    }

    $needConnect = $true
    if ($ctx) {
        $ctxScopes = @($ctx.Scopes)
        foreach ($candidate in $scopeCandidates) {
            if (Test-ContextSatisfiesScopes -ContextScopes $ctxScopes -RequiredScopes $candidate) {
                Write-Info ("Graph context already satisfies scopes: {0}" -f (($candidate -join ', ')))
                $needConnect = $false
                break
            }
        }
        if ($needConnect) {
            Write-Step "Existing Graph context is missing required scopes; attempting to reconnect..." Yellow
        }
    } else {
        Write-Step "No active Graph context detected; connecting..." Yellow
    }

    if ($needConnect) {
        [void](Connect-GraphWithFallback -ScopeSets $scopeCandidates)
        $selectCmd = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
        if ($selectCmd) {
            Select-MgProfile -Name beta -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

function Get-UserDetail {
    param([string]$UserId)

    $select = $UserSelectProperties -join ','
    $parameters = @{
        UserId   = $UserId
        Property = $UserSelectProperties
        ErrorAction = 'Stop'
    }

    try {
        $detail = Get-MgUser @parameters
        if (-not $detail) {
            throw "User '$UserId' could not be retrieved."
        }
        return $detail
    } catch {
        $message = $_.Exception.Message
        if ($message -match 'Request_ResourceNotFound' -or $message -match 'NotFound') {
            throw "User '$UserId' could not be found."
        }
        throw $_
    }
}

function Get-UserExtendedData {
    param([string]$UserId)

    $payload = [ordered]@{}

    try {
        $payload['LicenseDetails'] = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
    } catch {
        $payload['LicenseDetails'] = @()
    }

    try {
        $payload['AppRoleAssignments'] = Get-MgUserAppRoleAssignment -UserId $UserId -ErrorAction Stop
    } catch {
        $payload['AppRoleAssignments'] = @()
    }

    try {
        $payload['MemberOf'] = Get-MgUserMemberOf -UserId $UserId -ConsistencyLevel eventual -Top 20 -ErrorAction Stop
    } catch {
        $payload['MemberOf'] = @()
    }

    try {
        $payload['AuthenticationMethods'] = Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop
    } catch {
        $payload['AuthenticationMethods'] = @()
    }

    return [pscustomobject]$payload
}

function Build-NameString {
    param([object]$User)

    $surname = $User.Surname
    $givenName = $User.GivenName
    if ([string]::IsNullOrWhiteSpace($surname) -and [string]::IsNullOrWhiteSpace($givenName)) {
        return $User.DisplayName
    }
    $fullWidthSpace = [string][char]0x3000
    return "{0}{2}{1}" -f $surname, $givenName, $fullWidthSpace
}

function Get-WorksheetByName {
    param(
        [Parameter(Mandatory = $true)][object]$Workbook,
        [Parameter(Mandatory = $true)][string]$TargetName
    )

    $sheetNames = @()
    $count = $Workbook.Worksheets.Count
    $ordinalMatch = $null
    $ignoreCaseMatch = $null
    $prefixMatch = $null

    for ($i = 1; $i -le $count; $i++) {
        $name = [string]($Workbook.Worksheets.Item($i).Name)
        $sheetNames += $name
        if (-not $ordinalMatch -and [string]::Equals($name, $TargetName, [System.StringComparison]::Ordinal)) {
            $ordinalMatch = $i
        } elseif (-not $ignoreCaseMatch -and [string]::Equals($name, $TargetName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $ignoreCaseMatch = $i
        } elseif (-not $prefixMatch -and $name.StartsWith($TargetName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $prefixMatch = $i
        }
    }

    Write-Host ("[DEBUG] Worksheets detected: {0}" -f ($sheetNames -join ", ")) -ForegroundColor DarkGray

    $index = $ordinalMatch
    if (-not $index) { $index = $ignoreCaseMatch }
    if (-not $index) { $index = $prefixMatch }
    if ($index) {
        return $Workbook.Worksheets.Item($index)
    }

    throw "Worksheet '$TargetName' not found. Available: $($sheetNames -join ', ')"
}

function Write-RpaSheetValues {
    param(
        [string]$WorkbookPath,
        [hashtable]$CellMap
    )

    Write-Step "Writing user profile to RPA sheet..."
    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $sheet = Get-WorksheetByName -Workbook $workbook -TargetName 'RPAシート'

        foreach ($entry in $CellMap.GetEnumerator()) {
            $sheet.Range($entry.Key).Value2 = $entry.Value
        }

        $workbook.Save()
    } finally {
        if ($sheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
        if ($workbook) { $workbook.Close($true) }
        if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Read-RpaSheetValue {
    param(
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [Parameter(Mandatory = $true)][string]$Address
    )

    Write-Step ("Reading '{0}' from RPA sheet..." -f $Address)
    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $sheet = Get-WorksheetByName -Workbook $workbook -TargetName 'RPAシート'
        $value = $sheet.Range($Address).Value2
    } finally {
        if ($sheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
        if ($workbook) { $workbook.Close($true) }
        if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    return [string]$value
}

Write-Step ("Fetching user profile for {0}" -f $UserEmail)
Ensure-GraphModule
Connect-GraphIfNeeded -DesiredScopes $Scopes

$userDetail = Get-UserDetail -UserId $UserEmail
$extended = $null
if ($includeExtended) {
    $extended = Get-UserExtendedData -UserId $userDetail.Id
}

$workbookPath = Resolve-RpaBookPath
$nameValue = Build-NameString -User $userDetail
$company = $userDetail.CompanyName
$department = $userDetail.Department

$cellMap = @{
    'I5' = $nameValue
    'K5' = if ($company) { $company } else { '' }
    'L5' = if ($department) { $department } else { '' }
}
Write-RpaSheetValues -WorkbookPath $workbookPath -CellMap $cellMap

$secondaryEmailRaw = Read-RpaSheetValue -WorkbookPath $workbookPath -Address 'D3'
$secondaryEmail = if ($secondaryEmailRaw) { $secondaryEmailRaw.Trim() } else { '' }
$secondaryResult = $null
if ([string]::IsNullOrWhiteSpace($secondaryEmail)) {
    Write-Info "Secondary email (RPAシート!D3) is empty. Skipping secondary lookup."
} else {
    Write-Step ("Fetching secondary user profile for {0}" -f $secondaryEmail)
    try {
        $secondaryDetail = Get-UserDetail -UserId $secondaryEmail
        $secondaryName = Build-NameString -User $secondaryDetail
        Write-RpaSheetValues -WorkbookPath $workbookPath -CellMap @{ 'D15' = $secondaryName }
        $secondaryResult = [ordered]@{
            userEmail = $secondaryEmail
            nameFullWidth = $secondaryName
            userDetail = $secondaryDetail
        }
    } catch {
        Write-Warning ("Secondary user lookup failed: {0}" -f $_.Exception.Message)
    }
}

$output = [ordered]@{
    userEmail   = $UserEmail
    workbook    = $workbookPath
    userDetail  = $userDetail
    extended    = $extended
    nameFullWidth = $nameValue
}
if ($secondaryResult) {
    $output['secondaryUser'] = $secondaryResult
}

$output | ConvertTo-Json -Depth 10

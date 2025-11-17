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
    [int]$RequestTimeoutSeconds = 3,

    [Parameter(Mandatory = $false)]
    [int]$MaxDepth = 15
)

if ($Scopes.Count -eq 1 -and $Scopes[0] -match ',') {
    $Scopes = $Scopes[0].Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
}

$ErrorActionPreference = 'Stop'
$includeExtended = [bool]$IncludeExtendedData
$RequestTimeoutSeconds = [Math]::Max($RequestTimeoutSeconds, 1)
$MaxDepth = [Math]::Max($MaxDepth, 1)
$script:ScopeCandidates = $null

$ManagerSelectProperties = @(
    'id','displayName','mail','userPrincipalName','jobTitle','department','companyName',
    'businessPhones','mobilePhone','officeLocation','preferredLanguage','givenName','surname',
    'mailNickname','userType','accountEnabled','otherMails','imAddresses',
    'employeeId','employeeType','employeeHireDate','employeeOrgData','createdDateTime',
    'onPremisesSamAccountName','onPremisesUserPrincipalName','onPremisesDistinguishedName',
    'onPremisesDomainName','onPremisesImmutableId','country','city','state','postalCode','streetAddress',
    'usageLocation','preferredName','displayNamePronunciation'
)
$ManagerSelectQuery = ($ManagerSelectProperties -join ',')

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

function Write-DebugLog {
    param([string]$Message)
    Write-Host "[DEBUG] $Message" -ForegroundColor DarkYellow
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
            if (Test-GraphConsentError -Message $_.Exception.Message) {
                Write-Step ("Consent missing for scopes {0}: {1}" -f $label, $_.Exception.Message) Yellow
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
    param([string[]]$ContextScopes, [string[]]$RequiredScopes)

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

function Start-Timer {
    return [System.Diagnostics.Stopwatch]::StartNew()
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
                Write-Info ("Graph context already has scopes: {0}" -f (($candidate -join ', ')))
                $needConnect = $false
                break
            }
        }
        if ($needConnect) {
            Write-Step ("Existing context missing scopes; reconnecting...") Yellow
        } else {
            $needConnect = $false
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

function Get-UserManagerObject {
    param([string]$UserId)

    $totalTimer = Start-Timer
    $methodUsed = $null

    try {
        $managerCmd = Get-Command -Name Get-MgUserManagerUser -ErrorAction SilentlyContinue
        if ($managerCmd) {
            $callTimer = Start-Timer
            try {
                Write-DebugLog ("Manager lookup via Get-MgUserManagerUser started for {0}" -f $UserId)
                Write-Step ("Fetching manager via Get-MgUserManagerUser for {0}..." -f $UserId)
                $result = Get-MgUserManagerUser -UserId $UserId -ErrorAction Stop
                $callTimer.Stop()
                $methodUsed = 'Get-MgUserManagerUser'
                Write-Info ("Manager lookup via Get-MgUserManagerUser completed in {0} ms" -f $callTimer.ElapsedMilliseconds)
                return $result
            } catch {
                $callTimer.Stop()
                $message = $_.Exception.Message
                Write-Step "Get-MgUserManagerUser failed; trying alternate method..." Yellow
                Write-Info ("Get-MgUserManagerUser error after {0} ms: {1}" -f $callTimer.ElapsedMilliseconds, $message)
            }
        }

        $callTimer = Start-Timer
        Write-DebugLog ("Manager lookup via Get-MgUserManager started for {0}" -f $UserId)
        Write-Step ("Fetching manager via Get-MgUserManager for {0}..." -f $UserId)
        $mgr = Get-MgUserManager -UserId $UserId -ErrorAction Stop
        $callTimer.Stop()
        $methodUsed = 'Get-MgUserManager'
        Write-Info ("Manager lookup via Get-MgUserManager completed in {0} ms" -f $callTimer.ElapsedMilliseconds)
        if ($mgr -and ($mgr.PSObject.Properties.Name -contains 'AdditionalProperties') -and $mgr.AdditionalProperties) {
            $data = [ordered]@{}
            foreach ($entry in $mgr.AdditionalProperties.GetEnumerator()) {
                $data[$entry.Key] = $entry.Value
            }
            if ($mgr.PSObject.Properties.Name -contains 'Id' -and -not $data.Contains('id')) {
                $data['id'] = $mgr.Id
            }
            $result = [pscustomobject]$data
        } else {
            $result = $mgr
        }
        return $result
    } catch {
        $message = $_.Exception.Message
        if ($message -match 'Insufficient privileges' -or $message -match 'Authorization_RequestDenied') {
            throw 'Microsoft Graph permissions are insufficient. Request admin consent for User.Read.All and Directory.Read.All.'
        }
        if ($message -match 'Request_ResourceNotFound' -or $message -match 'ResourceNotFound' -or $message -match 'NotFound') {
            throw "User '$UserId' or their manager could not be found."
        }
        Write-Step ("Falling back to Invoke-MgGraphRequest for manager lookup of {0}..." -f $UserId) Yellow
        $uri = "/users/{0}/manager?$select=$ManagerSelectQuery" -f $UserId
        $callTimer = Start-Timer
        Write-DebugLog ("Manager lookup via Invoke-MgGraphRequest started for {0}" -f $UserId)
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop -RetryCount 1
        $callTimer.Stop()
        $methodUsed = 'Invoke-MgGraphRequest'
        Write-Info ("Manager lookup via Invoke-MgGraphRequest completed in {0} ms" -f $callTimer.ElapsedMilliseconds)
        return $result
    } finally {
        if ($totalTimer) {
            $totalTimer.Stop()
            $methodLabel = if ($methodUsed) { $methodUsed } else { 'unknown' }
            Write-Info ("Total manager lookup time for {0}: {1} ms (method={2})" -f $UserId, $totalTimer.ElapsedMilliseconds, $methodLabel)
        }
    }
}

function Get-ManagerDetail {
    param([object]$ManagerObject)

    if (-not $ManagerObject) { return $null }
    $managerId = $ManagerObject.Id
    if (-not $managerId) {
        if ($ManagerObject.PSObject.Properties.Name -contains 'userPrincipalName') {
            $managerId = $ManagerObject.userPrincipalName
        } elseif ($ManagerObject.PSObject.Properties.Name -contains 'mail') {
            $managerId = $ManagerObject.mail
        }
    }
    if (-not $managerId) { return $null }

    try {
        $callTimer = Start-Timer
        Write-DebugLog ("Manager detail lookup via Get-MgUser started for {0}" -f $managerId)
        Write-Step ("Fetching manager detail via Get-MgUser for {0}..." -f $managerId)
        $detail = Get-MgUser -UserId $managerId -Property $ManagerSelectProperties -ErrorAction Stop
        $callTimer.Stop()
        Write-Info ("Manager detail lookup via Get-MgUser completed in {0} ms" -f $callTimer.ElapsedMilliseconds)
        return $detail
    } catch {
        if ($callTimer) { $callTimer.Stop() }
        $message = $_.Exception.Message
        Write-Step ("Get-MgUser for manager failed; retrying with Invoke-MgGraphRequest for {0}..." -f $managerId) Yellow
        Write-Info ("Get-MgUser manager detail error: {0}" -f $message)
    }

    try {
        $uri = "/users/{0}?$select=$ManagerSelectQuery" -f $managerId
        $callTimer = Start-Timer
        Write-DebugLog ("Manager detail lookup via Invoke-MgGraphRequest started for {0}" -f $managerId)
        $detail = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop -RetryCount 1
        $callTimer.Stop()
        Write-Info ("Manager detail lookup via Invoke-MgGraphRequest completed in {0} ms" -f $callTimer.ElapsedMilliseconds)
        return $detail
    } catch {
        if ($callTimer) { $callTimer.Stop() }
        Write-Info ("Unable to fetch extended manager detail for {0}: {1}" -f $managerId, $_.Exception.Message)
        return $null
    }
}

function Get-UserExtendedData {
    param([string]$UserId)

    $payload = [ordered]@{}

    try {
        $timer = Start-Timer
        Write-DebugLog ("Extended data fetch (LicenseDetails) started for {0}" -f $UserId)
        $payload['LicenseDetails'] = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
        $timer.Stop()
        Write-Info ("Extended data fetch (LicenseDetails) completed in {0} ms" -f $timer.ElapsedMilliseconds)
    } catch {
        if ($timer) { $timer.Stop() }
        Write-Info ("Extended data fetch (LicenseDetails) failed for {0}: {1}" -f $UserId, $_.Exception.Message)
        $payload['LicenseDetails'] = @()
    }

    try {
        $timer = Start-Timer
        Write-DebugLog ("Extended data fetch (AppRoleAssignments) started for {0}" -f $UserId)
        $payload['AppRoleAssignments'] = Get-MgUserAppRoleAssignment -UserId $UserId -ErrorAction Stop
        $timer.Stop()
        Write-Info ("Extended data fetch (AppRoleAssignments) completed in {0} ms" -f $timer.ElapsedMilliseconds)
    } catch {
        if ($timer) { $timer.Stop() }
        Write-Info ("Extended data fetch (AppRoleAssignments) failed for {0}: {1}" -f $UserId, $_.Exception.Message)
        $payload['AppRoleAssignments'] = @()
    }

    try {
        $timer = Start-Timer
        Write-DebugLog ("Extended data fetch (MemberOf) started for {0}" -f $UserId)
        $payload['MemberOf'] = Get-MgUserMemberOf -UserId $UserId -ConsistencyLevel eventual -Top 20 -ErrorAction Stop
        $timer.Stop()
        Write-Info ("Extended data fetch (MemberOf) completed in {0} ms" -f $timer.ElapsedMilliseconds)
    } catch {
        if ($timer) { $timer.Stop() }
        Write-Info ("Extended data fetch (MemberOf) failed for {0}: {1}" -f $UserId, $_.Exception.Message)
        $payload['MemberOf'] = @()
    }

    try {
        $timer = Start-Timer
        Write-DebugLog ("Extended data fetch (AuthenticationMethods) started for {0}" -f $UserId)
        $payload['AuthenticationMethods'] = Get-MgUserAuthenticationMethod -UserId $UserId -ErrorAction Stop
        $timer.Stop()
        Write-Info ("Extended data fetch (AuthenticationMethods) completed in {0} ms" -f $timer.ElapsedMilliseconds)
    } catch {
        if ($timer) { $timer.Stop() }
        Write-Info ("Extended data fetch (AuthenticationMethods) failed for {0}: {1}" -f $UserId, $_.Exception.Message)
        $payload['AuthenticationMethods'] = @()
    }

    return [pscustomobject]$payload
}

function Build-NameString {
    param([object]$User)

    if (-not $User) { return $null }
    $surname = $User.Surname
    $givenName = $User.GivenName
    if ([string]::IsNullOrWhiteSpace($surname) -and [string]::IsNullOrWhiteSpace($givenName)) {
        if ($User.PSObject.Properties.Name -contains 'DisplayName') {
            return $User.DisplayName
        }
        if ($User.PSObject.Properties.Name -contains 'mail') {
            return $User.mail
        }
        return $User.userPrincipalName
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

    Write-DebugLog ("Worksheets detected: {0}" -f ($sheetNames -join ", "))

    $index = $ordinalMatch
    if (-not $index) { $index = $ignoreCaseMatch }
    if (-not $index) { $index = $prefixMatch }
    if ($index) {
        return $Workbook.Worksheets.Item($index)
    }

    throw "Worksheet '$TargetName' not found. Available: $($sheetNames -join ', ')"
}

function Open-Workbook {
    param([string]$WorkbookPath)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-DebugLog ("Requesting workbook open: {0}" -f $WorkbookPath)
    $timer = Start-Timer
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    $timer.Stop()
    Write-Info ("Workbook open completed in {0} ms" -f $timer.ElapsedMilliseconds)
    $sheet = Get-WorksheetByName -Workbook $workbook -TargetName 'RPAシート'

    return [pscustomobject]@{
        Excel    = $excel
        Workbook = $workbook
        Sheet    = $sheet
    }
}

function Close-Workbook {
    param($excelObj)

    if ($excelObj -ne $null) {
        Write-DebugLog "Closing workbook resources..."
        if ($excelObj.Sheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObj.Sheet) }
        if ($excelObj.Workbook) { $excelObj.Workbook.Close($true) }
        if ($excelObj.Excel) { $excelObj.Excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObj.Excel) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Info "Workbook resources released."
    }
}

Write-Step ("Building manager chain starting from {0}" -f $UserEmail)
Ensure-GraphModule
Connect-GraphIfNeeded -DesiredScopes $Scopes
if ($includeExtended) { Write-Info 'Extended manager data retrieval is enabled. This will slow down processing.' } else { Write-Info 'Extended manager data retrieval is disabled for faster execution.' }

$workbookPath = Resolve-RpaBookPath
Write-Info ("Target workbook path resolved: {0}" -f $workbookPath)
$excelObj = Open-Workbook -WorkbookPath $workbookPath
$results = @()
$visitedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

try {
    $currentId = $UserEmail
    for ($index = 0; $index -lt $MaxDepth; $index++) {
        $levelTimer = Start-Timer
        Write-DebugLog ("[Level {0}] Processing started for identifier {1}" -f $index, $currentId)
        Write-Step ("Resolving manager level {0} for identifier {1}" -f $index, $currentId)
        $lookupTimer = Start-Timer
        try {
            $managerObject = Get-UserManagerObject -UserId $currentId
        } catch {
            Write-Step ("Manager lookup failed at level {0}: {1}" -f $index, $_.Exception.Message) Yellow
            break
        }
        $lookupTimer.Stop()
        Write-Info ("Level {0} manager lookup completed in {1} ms" -f $index, $lookupTimer.ElapsedMilliseconds)

        if (-not $managerObject) {
            Write-Info ("No further manager object returned at level {0}; stopping chain." -f $index)
            break
        }

        $managerId = $managerObject.Id
        if (-not $managerId -and ($managerObject.PSObject.Properties.Name -contains 'userPrincipalName')) {
            $managerId = $managerObject.userPrincipalName
        }
        if (-not $managerId -and ($managerObject.PSObject.Properties.Name -contains 'mail')) {
            $managerId = $managerObject.mail
        }
        if ([string]::IsNullOrWhiteSpace($managerId)) {
            Write-Step ("Manager object at level {0} lacks an identifier; stopping chain." -f $index) Yellow
            break
        }

        if ($visitedIds.Contains($managerId)) {
            Write-Step ("Detected loop in manager chain at {0}; stopping to prevent infinite cycle." -f $managerId) Yellow
            break
        }
        $visitedIds.Add($managerId) | Out-Null

        $managerDetail = $null
        $detailElapsed = 0
        $detailTimer = Start-Timer
        try {
            $managerDetail = Get-ManagerDetail -ManagerObject $managerObject
        } catch {
            Write-Info ("Manager detail fetch failed for {0}: {1}" -f $managerId, $_.Exception.Message)
            $managerDetail = $null
        }
        finally {
            if ($detailTimer) { $detailTimer.Stop() }
        }
        $detailElapsed = $detailTimer.ElapsedMilliseconds
        if ($managerDetail) {
            Write-Info ("Level {0}: manager detail fetch took {1} ms" -f $index, $detailElapsed)
        } else {
            Write-DebugLog ("Level {0}: manager detail could not be retrieved; falling back to basic object.")
        }

        $detailForDisplay = if ($managerDetail) { $managerDetail } else { $managerObject }
        $displayName = Build-NameString -User $detailForDisplay
        $companyName = $null
        $department = $null
        $jobTitle = $null
        $userPrincipalName = $null

        if ($managerDetail) {
            $companyName = $managerDetail.CompanyName
            $department = $managerDetail.Department
            $jobTitle = $managerDetail.JobTitle
            if ($managerDetail.PSObject.Properties.Name -contains 'UserPrincipalName') {
                $userPrincipalName = $managerDetail.UserPrincipalName
            }
        } else {
            if ($managerObject.PSObject.Properties.Name -contains 'companyName') {
                $companyName = $managerObject.companyName
            }
            if ($managerObject.PSObject.Properties.Name -contains 'department') {
                $department = $managerObject.department
            }
            if ($managerObject.PSObject.Properties.Name -contains 'jobTitle') {
                $jobTitle = $managerObject.jobTitle
            }
            if ($managerObject.PSObject.Properties.Name -contains 'userPrincipalName') {
                $userPrincipalName = $managerObject.userPrincipalName
            }
        }

        $extended = $null
        $extendedElapsed = 0
        if ($includeExtended) {
            $extendedTimer = Start-Timer
            try {
                $extended = Get-UserExtendedData -UserId $managerId
            } catch {
                Write-Info ("Extended data fetch failed for {0}: {1}" -f $managerId, $_.Exception.Message)
                $extended = $null
            }
            $extendedTimer.Stop()
            $extendedElapsed = $extendedTimer.ElapsedMilliseconds
            Write-Info ("Level {0}: extended data fetch took {1} ms" -f $index, $extendedElapsed)
        }
        else {
            Write-DebugLog ("Level {0}: skipping extended data fetch (IncludeExtendedData disabled)" -f $index)
        }

        $row = 6 + $index
        $sheet = $excelObj.Sheet
        $mailAddress = if ($managerDetail -and $managerDetail.Mail) { $managerDetail.Mail } elseif ($managerObject.PSObject.Properties.Name -contains 'mail') { $managerObject.mail } else { '' }

        $sheet.Cells.Item($row, 9).Value2  = $displayName
        $sheet.Cells.Item($row, 10).Value2 = $mailAddress
        $sheet.Cells.Item($row, 11).Value2 = $companyName
        $sheet.Cells.Item($row, 12).Value2 = $department

        Write-Info ("Level {0} processing finished in {1} ms" -f $index, $lookupTimer.ElapsedMilliseconds + ($detailTimer?.ElapsedMilliseconds ?? 0) + ($extendedTimer?.ElapsedMilliseconds ?? 0))
        $results += [pscustomobject]@{
            Index          = $index
            Identifier     = $managerId
            DisplayName    = $displayName
            Mail           = $mailAddress
            JobTitle       = $jobTitle
            CompanyName    = $companyName
            Department     = $department
            Detail         = $managerDetail
            RawObject      = $managerObject
            Extended       = $extended
        }

        if ($managerDetail -and $managerDetail.UserPrincipalName) {
            $currentId = $managerDetail.UserPrincipalName
        } else {
            $currentId = $managerId
        }

        $levelTimer.Stop()
        Write-Info ("Level {0} total elapsed {1} ms (lookup={2} ms, detail={3} ms, extended={4} ms)" -f $index, $levelTimer.ElapsedMilliseconds, $lookupTimer.ElapsedMilliseconds, $detailElapsed, $extendedElapsed)
    }
}
finally {
    Close-Workbook -excelObj $excelObj
}

$output = [ordered]@{
    userEmail     = $UserEmail
    workbook      = $workbookPath
    managerCount  = $results.Count
    managers      = $results
}

$output | ConvertTo-Json -Depth 10

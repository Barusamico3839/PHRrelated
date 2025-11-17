param(
    [Parameter(Mandatory = $false)]
    [string[]]$Scopes = @('User.Read.All','Directory.Read.All'),

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceAuth,

    [Parameter(Mandatory = $false)]
    [switch]$UseBrowserAuth,

    [Parameter(Mandatory = $false)]
    [switch]$SkipModuleInstall,

    [Parameter(Mandatory = $false)]
    [int]$RequestTimeoutSeconds = 3
)

if ($Scopes.Count -eq 1 -and $Scopes[0] -match ',') {
    $Scopes = $Scopes[0].Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
}

$ErrorActionPreference = 'Stop'
$script:DesiredScopes = $Scopes
$script:SkipModuleInstall = $SkipModuleInstall
$script:RequestTimeoutSeconds = [Math]::Max($RequestTimeoutSeconds, 1)

if ($PSBoundParameters.ContainsKey('UseDeviceAuth')) {
    $script:UseDeviceAuth = [bool]$UseDeviceAuth
} elseif ($PSBoundParameters.ContainsKey('UseBrowserAuth')) {
    $script:UseDeviceAuth = -not [bool]$UseBrowserAuth
} else {
    # Default: always use device authentication unless explicitly overridden.
    $script:UseDeviceAuth = $true
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
    param(
        [System.Collections.Generic.List[object]]$ScopeSets,
        [bool]$UseDeviceAuth = $false
    )

    if (-not $ScopeSets -or $ScopeSets.Count -eq 0) {
        throw 'No Graph scope sets specified.'
    }

    $lastError = $null
    foreach ($scopeSet in $ScopeSets) {
        $scopes = @($scopeSet)
        $label = if ($scopes.Count -gt 0) { $scopes -join ', ' } else { '(default)' }
        Write-Host ("[INFO] Attempting Graph sign-in with scopes: {0}" -f $label) -ForegroundColor DarkGray
        try {
            if ($UseDeviceAuth) {
                Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication | Out-Null
            } else {
                Connect-MgGraph -Scopes $scopes | Out-Null
            }
            Write-Host ("[INFO] Graph sign-in succeeded with scopes: {0}" -f $label) -ForegroundColor DarkGreen
            return $scopes
        } catch [System.Management.Automation.PipelineStoppedException] {
            throw
        } catch {
            $lastError = $_
            $message = $_.Exception.Message
            if (Test-GraphConsentError -Message $message) {
                Write-Host ("[WARNING] Consent missing for scopes {0}: {1}" -f $label, $message) -ForegroundColor Yellow
                continue
            }
            throw
        }
    }

    if ($lastError) {
        throw $lastError
    }
    throw 'Graph sign-in failed for all scope sets.'
}

function Resolve-RpaBookPath {
    return [System.IO.Path]::Combine(
        $env:USERPROFILE,
        'Desktop',
        '【全社標準】弔事対応フォルダ',
        '2. RPAブック',
        'RPAブック.xlsx'
    )
}

function Get-MailFromWorkbook {
    param([string]$WorkbookPath)

    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
        throw "Workbook not found: $WorkbookPath"
    }

    $excel = $null
    $workbook = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)
        $sheet = $workbook.Worksheets.Item('RPAシート')
        return [string]($sheet.Range('J5').Text).Trim()
    } finally {
        if ($sheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Set-GraphTimeout {
    $timeoutCmd = Get-Command -Name Set-MgGraphRequestTimeout -ErrorAction SilentlyContinue
    if ($timeoutCmd) {
        try {
            Set-MgGraphRequestTimeout -Milliseconds ($script:RequestTimeoutSeconds * 1000)
        } catch {
            # ignore timeout failures
        }
    }
}

function Ensure-GraphModule {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if (-not $script:SkipModuleInstall) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue | Out-Null
    Set-GraphTimeout
    $sw.Stop()
    Write-Host ("[INFO] Ensure-GraphModule took {0} ms" -f $sw.ElapsedMilliseconds) -ForegroundColor DarkGray
}

Ensure-GraphModule

if ($script:UseDeviceAuth) {
    Write-Host "[STEP] デバイスコード認証フローを開始します。" -ForegroundColor Cyan
    Write-Host "       表示されたコードを https://microsoft.com/devicelogin に入力し、認証を完了してください。" -ForegroundColor DarkCyan
} else {
    Write-Host "[STEP] ブラウザを使用した認証フローを開始します。" -ForegroundColor Cyan
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
$scopeCandidates = Get-GraphScopeCandidates -PrimaryScopes $script:DesiredScopes
try {
    $activeScopes = Connect-GraphWithFallback -ScopeSets $scopeCandidates -UseDeviceAuth:$script:UseDeviceAuth
} catch [System.Management.Automation.PipelineStoppedException] {
    throw 'Graph sign-in was interrupted (PipelineStoppedException). Re-run and complete the sign-in prompt, or retry with -UseDeviceAuth.'
}

$selectCmd = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
if ($selectCmd) {
    Select-MgProfile -Name beta -ErrorAction SilentlyContinue | Out-Null
}
$ctx = Get-MgContext -ErrorAction SilentlyContinue
Write-Host "[STEP] Graph authentication completed." -ForegroundColor DarkGreen

Write-Host "[STEP] Loading RPA workbook (J5) to retrieve target mail..." -ForegroundColor Cyan
$workbookPath = Resolve-RpaBookPath
$mailHonnin = Get-MailFromWorkbook -WorkbookPath $workbookPath

if (-not $mailHonnin) {
    throw "RPAシート J5 からメールアドレスを取得できませんでした。"
}

Write-Host ("[STEP] Target mail address detected: {0}" -f $mailHonnin) -ForegroundColor DarkCyan

$result = [ordered]@{
    mail_honnin      = $mailHonnin
    graph_connected  = [bool]$ctx
    active_scopes    = if ($ctx) { @($ctx.Scopes) } elseif ($activeScopes) { @($activeScopes) } else { @() }
}

$result | ConvertTo-Json -Depth 5

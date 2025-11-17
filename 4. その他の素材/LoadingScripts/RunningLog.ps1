Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore

# ==== XAML Load ====
$xamlPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "MyWindow.xaml"
$xamlText = Get-Content $xamlPath -Raw -Encoding UTF8
$xamlText = $xamlText.TrimStart([char]0xFEFF, [char]0x200B, [char]0x00)
$window   = [Windows.Markup.XamlReader]::Parse($xamlText)

# ==== Place window at top right ====
$screenWidth  = [System.Windows.SystemParameters]::PrimaryScreenWidth
$window.Left  = $screenWidth - $window.Width
$window.Top   = 0

# ==== Get controls ====
$loadingImage = $window.FindName("loadingImage")
$btnStop      = $window.FindName("btnStop")

# ==== Frames list ====
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$frames = @(
    (Join-Path $scriptRoot "frame1.png"),
    (Join-Path $scriptRoot "frame2.png"),
    (Join-Path $scriptRoot "frame3.png"),
    (Join-Path $scriptRoot "frame4.png")
)

function Get-BitmapImage($path) {
    if (-not (Test-Path $path)) { return $null }
    $bi = New-Object Windows.Media.Imaging.BitmapImage
    $bi.BeginInit()
    $bi.UriSource = (New-Object System.Uri (Resolve-Path $path))
    $bi.CacheOption = [Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bi.CreateOptions = [Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache
    $bi.EndInit()
    $bi.Freeze()
    return $bi
}

# ==== Stop flag ====
$script:running = $true

# ==== Button event (just close window) ====
$btnStop.Add_Click({
    $script:running = $false
    Write-Host "Window closed by user."
    $window.Close()
})

# ==== Show first frame ====
$bmp = Get-BitmapImage $frames[0]
if ($bmp) {
    $loadingImage.Source = $bmp
    Write-Host "Frame 1 displayed."
}

# ==== DispatcherTimer for animation ====
$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = New-Object TimeSpan 0,0,0,0,700  # 0.7 sec

$script:index = 0

$timer.add_Tick({
    if (-not $script:running) {
        $timer.Stop()
        return
    }

    $bmp = Get-BitmapImage $frames[$script:index]
    if ($bmp) {
        $loadingImage.Source = $bmp
        Write-Host ("Frame displayed: {0} ({1})" -f ($script:index + 1), (Split-Path $frames[$script:index] -Leaf))
    }

    # Move to next frame (loop)
    $script:index = ($script:index + 1) % $frames.Count
})

$timer.Start()

# ==== Show window ====
$window.ShowDialog() | Out-Null

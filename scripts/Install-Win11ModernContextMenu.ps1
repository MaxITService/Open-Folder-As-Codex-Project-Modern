#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    [ValidateSet('x64', 'ARM64')]
    [string] $Platform = 'x64',

    [string] $CodexExe,
    [switch] $Uninstall,
    [switch] $NoRestartExplorer
)

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$modernRoot = Split-Path -Parent $scriptRoot
$artifactsRoot = Join-Path $modernRoot 'artifacts'
$packageRoot = Join-Path $artifactsRoot 'PackageRoot'
$packageAssets = Join-Path $packageRoot 'Assets'
$certRoot = Join-Path $artifactsRoot 'cert'
$msixOut = Join-Path $artifactsRoot 'OpenFolderAsCodexProject.Win11Modern.msix'
$pfxPath = Join-Path $certRoot 'OpenFolderAsCodexProject.Dev.pfx'
$cerPath = Join-Path $certRoot 'OpenFolderAsCodexProject.Dev.cer'
$publisher = 'CN=OpenFolderAsCodexProjectDev'
$packageName = 'OpenFolderAsCodexProject.Win11Modern'
$version = '0.1.0.0'
$clsid = '5f220f1e-376f-4f5b-9d5e-d42d924ff811'
$appId = '2d737260-9d21-44d1-89a7-e50f438da3c3'
$settingsKey = 'HKCU:\Software\OpenFolderAsCodexProject\Win11Modern'
$pfxPassword = ConvertTo-SecureString -String 'OpenFolderAsCodexProject-Dev' -Force -AsPlainText
$codexInstallUrl = 'https://apps.microsoft.com/detail/9plm9xgg6vks'
$buildToolsUrl = 'https://visualstudio.microsoft.com/visual-cpp-build-tools/'
$windowsSdkUrl = 'https://developer.microsoft.com/windows/downloads/windows-sdk/'

function Test-IsAdmin {
    $principal = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-CodexDesktopExe {
    param([string] $ExplicitPath)

    if ($ExplicitPath) {
        $resolved = Resolve-Path -LiteralPath $ExplicitPath -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $resolved.Path -PathType Leaf)) {
            throw "Codex executable is not a file: $ExplicitPath"
        }
        return $resolved.Path
    }

    $package = Get-AppxPackage -Name OpenAI.Codex -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($package) {
        $candidate = Join-Path $package.InstallLocation 'app\Codex.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }
    }

    $localCandidate = Join-Path $env:LOCALAPPDATA 'OpenAI\Codex\app\Codex.exe'
    if (Test-Path -LiteralPath $localCandidate -PathType Leaf) {
        return $localCandidate
    }

    throw 'Could not find Codex Desktop. Pass -CodexExe "C:\Path\To\Codex.exe".'
}

function Find-VsWhere {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe",
        "$env:ProgramFiles\Microsoft Visual Studio\Installer\vswhere.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }

    throw 'Could not find vswhere.exe. Install Visual Studio 2022 Build Tools with the C++ desktop workload.'
}

function Find-MSBuild {
    $vswhere = Find-VsWhere
    $path = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -find 'MSBuild\Current\Bin\MSBuild.exe' | Select-Object -First 1
    if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw 'Could not find MSBuild.exe with the MSVC C++ toolset. Install Visual Studio Build Tools for C++.'
    }
    return $path
}

function Find-WindowsSdkTool {
    param([string] $ToolName)

    $root = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        throw 'Could not find Windows SDK bin folder. Install the Windows 10/11 SDK.'
    }

    $tool = Get-ChildItem -LiteralPath $root -Directory |
        Sort-Object Name -Descending |
        ForEach-Object { Join-Path $_.FullName "x64\$ToolName" } |
        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Select-Object -First 1

    if (-not $tool) {
        throw "Could not find $ToolName in the Windows SDK."
    }
    return $tool
}

function Test-Dependencies {
    param([string] $ExplicitCodexExe)

    Write-Host 'Checking dependencies...'
    $missing = New-Object System.Collections.Generic.List[string]

    try {
        $codexPath = Get-CodexDesktopExe -ExplicitPath $ExplicitCodexExe
        Write-Host "  OK Codex Desktop: $codexPath"
    }
    catch {
        $missing.Add('Codex Desktop')
    }

    try {
        $msbuild = Find-MSBuild
        Write-Host "  OK Visual Studio C++ Build Tools: $msbuild"
    }
    catch {
        $missing.Add('Visual Studio C++ Build Tools')
    }

    $missingSdk = $false
    try {
        $makeAppx = Find-WindowsSdkTool -ToolName 'MakeAppx.exe'
        Write-Host "  OK MakeAppx.exe: $makeAppx"
    }
    catch {
        $missingSdk = $true
    }

    try {
        $signTool = Find-WindowsSdkTool -ToolName 'SignTool.exe'
        Write-Host "  OK SignTool.exe: $signTool"
    }
    catch {
        $missingSdk = $true
    }

    if ($missingSdk) {
        $missing.Add('Windows SDK')
    }

    if ($missing.Count -eq 0) {
        Write-Host 'Dependency check passed.'
        return
    }

    Write-Host ''
    Write-Host 'Missing required dependencies:'
    foreach ($item in $missing) {
        Write-Host "  - $item"
    }

    Write-Host ''
    Write-Host 'Install links:'
    if ($missing.Contains('Codex Desktop')) {
        Write-Host "  Codex Desktop: $codexInstallUrl"
    }
    if ($missing.Contains('Visual Studio C++ Build Tools')) {
        Write-Host "  Visual Studio Build Tools for C++: $buildToolsUrl"
        Write-Host '    Select the C++ build tools / Desktop development with C++ components.'
    }
    if ($missing.Contains('Windows SDK')) {
        Write-Host "  Windows SDK: $windowsSdkUrl"
        Write-Host '    The SDK provides MakeAppx.exe and SignTool.exe.'
    }

    throw 'Install the missing dependencies and run this installer again.'
}

function Ensure-DevCertificate {
    New-Item -ItemType Directory -Force -Path $certRoot | Out-Null

    $cert = Get-ChildItem Cert:\LocalMachine\My |
        Where-Object { $_.Subject -eq $publisher -and $_.HasPrivateKey } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $cert) {
        Write-Host "Creating development certificate: $publisher"
        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $publisher `
            -CertStoreLocation Cert:\LocalMachine\My `
            -KeyExportPolicy Exportable `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears(3)
    }

    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword -Force | Out-Null

    $trusted = Get-ChildItem Cert:\LocalMachine\Root |
        Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
        Select-Object -First 1

    if (-not $trusted) {
        Write-Host 'Trusting development certificate in LocalMachine\Root'
        Import-Certificate -FilePath $cerPath -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
    }
}

function Write-PlaceholderPng {
    param(
        [string] $Path,
        [int] $Size
    )

    Add-Type -AssemblyName System.Drawing
    $bitmap = New-Object Drawing.Bitmap $Size, $Size
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.Clear([Drawing.Color]::FromArgb(16, 124, 16))
        $font = New-Object Drawing.Font 'Segoe UI', ([Math]::Max(8, [int]($Size / 3))), ([Drawing.FontStyle]::Bold), ([Drawing.GraphicsUnit]::Pixel)
        $brush = [Drawing.Brushes]::White
        $format = New-Object Drawing.StringFormat
        $format.Alignment = [Drawing.StringAlignment]::Center
        $format.LineAlignment = [Drawing.StringAlignment]::Center
        $rectangle = New-Object Drawing.RectangleF -ArgumentList 0, 0, $Size, $Size
        $graphics.DrawString('C', $font, $brush, $rectangle, $format)
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Render-Manifest {
    $architecture = switch ($Platform) {
        'x64' { 'x64' }
        'ARM64' { 'arm64' }
    }

    $template = Get-Content -Raw -LiteralPath (Join-Path $modernRoot 'package\AppxManifest.xml.template')
    $manifest = $template.
        Replace('{{PACKAGE_NAME}}', $packageName).
        Replace('{{PUBLISHER}}', $publisher).
        Replace('{{VERSION}}', $version).
        Replace('{{ARCHITECTURE}}', $architecture).
        Replace('{{CLSID}}', $clsid).
        Replace('{{APP_ID}}', $appId)

    Set-Content -LiteralPath (Join-Path $packageRoot 'AppxManifest.xml') -Value $manifest -Encoding UTF8
}

function Build-Dll {
    $msbuild = Find-MSBuild
    $project = Join-Path $modernRoot 'src\OpenFolderAsCodexCommand.vcxproj'
    & $msbuild $project /m /restore /p:Configuration=$Configuration /p:Platform=$Platform | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed with exit code $LASTEXITCODE."
    }

    $dll = Join-Path $artifactsRoot "bin\$Platform\$Configuration\OpenFolderAsCodexCommand.dll"
    if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) {
        throw "Build completed but DLL was not found: $dll"
    }
    return $dll
}

function Build-Host {
    $msbuild = Find-MSBuild
    $project = Join-Path $modernRoot 'src\CodexContextMenuHost.vcxproj'
    & $msbuild $project /m /restore /p:Configuration=$Configuration /p:Platform=$Platform | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Host MSBuild failed with exit code $LASTEXITCODE."
    }

    $hostExe = Join-Path $artifactsRoot "bin\$Platform\$Configuration\CodexContextMenuHost.exe"
    if (-not (Test-Path -LiteralPath $hostExe -PathType Leaf)) {
        throw "Build completed but host exe was not found: $hostExe"
    }
    return $hostExe
}

function Install-Package {
    param(
        [string] $DllPath,
        [string] $HostPath
    )

    $makeAppx = Find-WindowsSdkTool -ToolName 'MakeAppx.exe'
    $signTool = Find-WindowsSdkTool -ToolName 'SignTool.exe'

    if (Test-Path -LiteralPath $packageRoot) {
        Remove-Item -LiteralPath $packageRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $packageRoot, $packageAssets | Out-Null

    Copy-Item -LiteralPath $DllPath -Destination (Join-Path $packageRoot 'OpenFolderAsCodexCommand.dll') -Force
    Copy-Item -LiteralPath $HostPath -Destination (Join-Path $packageRoot 'CodexContextMenuHost.exe') -Force
    Write-PlaceholderPng -Path (Join-Path $packageAssets 'Square44x44Logo.png') -Size 44
    Write-PlaceholderPng -Path (Join-Path $packageAssets 'Square150x150Logo.png') -Size 150
    Write-PlaceholderPng -Path (Join-Path $packageAssets 'StoreLogo.png') -Size 50
    Render-Manifest

    if (Test-Path -LiteralPath $msixOut -PathType Leaf) {
        Remove-Item -LiteralPath $msixOut -Force
    }

    & $makeAppx pack /d $packageRoot /p $msixOut /o
    if ($LASTEXITCODE -ne 0) {
        throw "MakeAppx failed with exit code $LASTEXITCODE."
    }

    & $signTool sign /fd SHA256 /f $pfxPath /p 'OpenFolderAsCodexProject-Dev' $msixOut
    if ($LASTEXITCODE -ne 0) {
        throw "SignTool failed with exit code $LASTEXITCODE."
    }

    Add-AppxPackage -Path $msixOut -ForceApplicationShutdown
}

function Restart-Explorer {
    if ($NoRestartExplorer) {
        return
    }

    Write-Host 'Restarting Explorer to refresh context menu registration.'
    Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Process explorer.exe
}

function Uninstall-ModernContextMenuPackage {
    $packages = Get-AppxPackage -Name $packageName -AllUsers -ErrorAction SilentlyContinue
    if ($packages) {
        $packages | Remove-AppxPackage -AllUsers
    }

    Remove-Item -LiteralPath $settingsKey -Recurse -Force -ErrorAction SilentlyContinue
    Restart-Explorer
    Write-Host "Uninstalled $packageName."
}

if (-not (Test-IsAdmin)) {
    throw 'This installer must be run elevated. It trusts a dev certificate and installs a Windows package.'
}

if ($Uninstall) {
    Uninstall-ModernContextMenuPackage
    return
}

$installedPackage = Get-AppxPackage -Name $packageName -AllUsers -ErrorAction SilentlyContinue | Select-Object -First 1
if ($installedPackage) {
    Write-Host "$packageName is already installed."
    Write-Host 'Y = uninstall it now, R = reinstall/update it, N = cancel'
    $answer = Read-Host 'What do you want to do? [Y/R/N]'

    switch -Regex ($answer) {
        '^[Yy]' {
            Uninstall-ModernContextMenuPackage
            return
        }
        '^[Rr]' {
            Write-Host 'Reinstalling package...'
        }
        default {
            Write-Host 'Cancelled. Nothing changed.'
            return
        }
    }
}

Test-Dependencies -ExplicitCodexExe $CodexExe

$codexDesktopExe = Get-CodexDesktopExe -ExplicitPath $CodexExe
New-Item -ItemType Directory -Force -Path $artifactsRoot | Out-Null
New-Item -Path $settingsKey -Force | Out-Null
New-ItemProperty -Path $settingsKey -Name 'CodexExe' -Value $codexDesktopExe -PropertyType ExpandString -Force | Out-Null
New-ItemProperty -Path $settingsKey -Name 'Title' -Value 'Open project in Codex' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $settingsKey -Name 'Enabled' -Value '1' -PropertyType String -Force | Out-Null

Ensure-DevCertificate
$dll = Build-Dll
$hostExe = Build-Host
Install-Package -DllPath $dll -HostPath $hostExe
Restart-Explorer

Write-Host "Installed Windows 11 modern context menu package: $packageName"
Write-Host "Codex Desktop: $codexDesktopExe"
Write-Host "Package root: $packageRoot"

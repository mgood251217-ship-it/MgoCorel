param(
    [string]$Version = "CorelDRAW Graphics Suite 2022",
    [switch]$Uninstall,
    [switch]$Elevated
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot

function Stop-Script {
    param([int]$Code)

    if ($Elevated) {
        Read-Host "Tekan Enter untuk menutup" | Out-Null
    }

    exit $Code
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Meminta hak Administrator..." -ForegroundColor Yellow

    $argList = @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-Version", "`"$Version`"",
        "-Elevated"
    )

    if ($Uninstall) {
        $argList += "-Uninstall"
    }

    Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Verb RunAs -Wait
    exit 0
}

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "       MgoCorel Addon Installer" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

try {
    $corelRoot = Join-Path $env:ProgramFiles "Corel"
    $suiteDir = Join-Path $corelRoot $Version
    $addonsDir = Join-Path $suiteDir "Programs64\Addons"
    $dest = Join-Path $addonsDir "MgoCorel"
    $staleDest = Join-Path $suiteDir "Draw\Addons\MgoCorel"

    if (-not (Test-Path $addonsDir)) {
        Write-Host "Folder tidak ditemukan: $addonsDir" -ForegroundColor Red
        Write-Host "Versi yang terpasang:" -ForegroundColor Yellow

        Get-ChildItem $corelRoot -Directory -ErrorAction SilentlyContinue |
            ForEach-Object { Write-Host "  $($_.Name)" }

        Stop-Script 1
    }

    if (Test-Path $staleDest) {
        Remove-Item $staleDest -Recurse -Force
        Write-Host "[OK] Folder lama dihapus: $staleDest"
    }

    if ($Uninstall) {
        if (Test-Path $dest) {
            Remove-Item $dest -Recurse -Force
        }

        Write-Host "Addon dihapus: $dest" -ForegroundColor Green
        Stop-Script 0
    }

    $addonSource = Join-Path $Root "addon"
    $iconSource = Join-Path $Root "src\icons"
    $addonMarker = Join-Path $addonSource "Coreldrw.addon"
    $appUi = Join-Path $addonSource "AppUI.xslt"
    $userUi = Join-Path $addonSource "UserUI.xslt"

    foreach ($required in @($addonMarker, $appUi, $userUi)) {
        if (-not (Test-Path $required)) {
            Write-Host "File wajib tidak ditemukan: $required" -ForegroundColor Red
            Stop-Script 1
        }
    }

    foreach ($optional in @("Resources.dll", "config.xml")) {
        if (-not (Test-Path (Join-Path $addonSource $optional))) {
            Write-Host "[WARN] $optional tidak ada. Jalankan task Build MgoCorel Addon dulu, jika tidak ikon tidak akan muncul." -ForegroundColor Yellow
        }
    }

    if (-not (Select-String -Path $appUi -Pattern "Corel Framework Data" -Quiet)) {
        Write-Host "AppUI.xslt tidak memakai namespace Corel Framework Data. Instalasi dibatalkan." -ForegroundColor Red
        Stop-Script 1
    }

    if (Test-Path $dest) {
        Remove-Item $dest -Recurse -Force
    }

    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    Get-ChildItem $addonSource -File | ForEach-Object {
        Copy-Item $_.FullName $dest -Force
        Write-Host "[OK] $($_.Name)"
    }

    if (Test-Path $iconSource) {
        $iconDest = Join-Path $dest "icons"
        New-Item -ItemType Directory -Path $iconDest -Force | Out-Null

        Get-ChildItem $iconSource -Filter "*.ico" -File | ForEach-Object {
            Copy-Item $_.FullName $iconDest -Force
            Write-Host "[OK] Icon: $($_.Name)"
        }
    }

    Write-Host ""
    Write-Host "Terpasang di: $dest" -ForegroundColor Green
    Write-Host "Tutup lalu buka ulang CorelDRAW."
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "[FAIL] $($_.Exception.Message)" -ForegroundColor Red
    Stop-Script 1
}

Stop-Script 0
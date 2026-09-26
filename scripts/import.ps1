$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $PSScriptRoot
$ModulesPath = Join-Path $Root "src\modules"
$FormsPath = Join-Path $Root "src\forms"

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "          MgoCorel Build" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

$corel = $null
$createdCorel = $false

try {
    $corel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("CorelDRAW.Application.24")
    Write-Host "CorelDRAW ditemukan." -ForegroundColor Green
}
catch {
    Write-Host "CorelDRAW belum berjalan, membuat instance..." -ForegroundColor Yellow
    $corel = New-Object -ComObject "CorelDRAW.Application.24"
    $corel.Visible = $true
    $createdCorel = $true
}

Write-Host "Version: $($corel.Version)"

Write-Host ""
Write-Host "Menginisialisasi VBA..." -ForegroundColor Cyan

try {
    $corel.InitializeVBA()
}
catch {
}

Start-Sleep -Milliseconds 500

Write-Host "Mencari VBA Project MgoCorel..." -ForegroundColor Cyan

$vbe = $corel.VBE
$vbProjects = $vbe.VBProjects
$vbProject = $null

for ($i = 1; $i -le $vbProjects.Count; $i++) {
    $project = $vbProjects.Item($i)

    if ($project.Name -eq "MgoCorel") {
        $vbProject = $project
        break
    }
}

if ($null -eq $vbProject) {
    Write-Host "Project MgoCorel tidak ditemukan di VBE." -ForegroundColor Red
    exit 1
}

Write-Host "Project MgoCorel ditemukan." -ForegroundColor Green

$gmsManager = $corel.GMSManager
$gmsProjects = $gmsManager.Projects
$gmsProject = $null

for ($i = 1; $i -le $gmsProjects.Count; $i++) {
    $project = $gmsProjects.Item($i)

    if ($project.Name -eq "MgoCorel") {
        $gmsProject = $project
        break
    }
}

if ($null -eq $gmsProject) {
    Write-Host "GMS MgoCorel tidak ditemukan." -ForegroundColor Red
    exit 1
}

Write-Host "GMS: $($gmsProject.FullFileName)" -ForegroundColor Green

$iconSource = Join-Path $PSScriptRoot "..\src\icons"
$iconSource = [System.IO.Path]::GetFullPath($iconSource)

if (Test-Path $iconSource) {
    $gmsUserPath = $gmsManager.UserGMSPath
    $iconDestination = Join-Path $gmsUserPath "icons"

    if (-not (Test-Path $iconDestination)) {
        New-Item -ItemType Directory -Path $iconDestination -Force | Out-Null
    }

    $icons = Get-ChildItem -Path $iconSource -Filter "*.ico" -File

    foreach ($icon in $icons) {
        $destination = Join-Path $iconDestination $icon.Name
        Copy-Item -Path $icon.FullName -Destination $destination -Force
        Write-Host "[OK] Icon: $($icon.Name)"
    }
}
else {
    Write-Host "[INFO] Folder icons tidak ditemukan: $iconSource" -ForegroundColor Yellow
}

$sourceFiles = @()

if (Test-Path $ModulesPath) {
    $sourceFiles += Get-ChildItem $ModulesPath -File |
        Where-Object {
            $_.Extension -in ".bas", ".cls"
        }
}

if (Test-Path $FormsPath) {
    $sourceFiles += Get-ChildItem $FormsPath -File -Filter "*.frm"
}

if ($sourceFiles.Count -eq 0) {
    Write-Host ""
    Write-Host "Tidak ada source file ditemukan." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Source yang ditemukan:" -ForegroundColor Cyan

foreach ($file in $sourceFiles) {
    Write-Host "  $($file.Name)"
}

Write-Host ""
Write-Host "Menghapus component yang dikelola source..." -ForegroundColor Cyan

$components = $vbProject.VBComponents

foreach ($file in $sourceFiles) {
    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    try {
        $component = $components.Item($componentName)

        if ($null -ne $component) {
            Write-Host "  Remove: $componentName" -ForegroundColor DarkYellow
            $components.Remove($component)
        }
    }
    catch {
    }
}

Start-Sleep -Milliseconds 300

$components = $vbProject.VBComponents

Write-Host ""
Write-Host "Import modules..." -ForegroundColor Cyan

$moduleFiles = $sourceFiles |
    Where-Object {
        $_.Extension -in ".bas", ".cls"
    }

foreach ($file in $moduleFiles) {
    Write-Host "  Import: $($file.Name)" -ForegroundColor Green

    try {
        $components.Import($file.FullName) | Out-Null
    }
    catch {
        Write-Host "  [FAIL] $($file.Name)" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "Import forms..." -ForegroundColor Cyan

$formFiles = $sourceFiles |
    Where-Object {
        $_.Extension -eq ".frm"
    }

foreach ($file in $formFiles) {
    Write-Host "  Import: $($file.Name)" -ForegroundColor Green

    try {
        $components.Import($file.FullName) | Out-Null
    }
    catch {
        Write-Host "  [FAIL] $($file.Name)" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

Start-Sleep -Milliseconds 500

Write-Host ""
Write-Host "Verifikasi component..." -ForegroundColor Cyan

$components = $vbProject.VBComponents
$failedComponents = @()

foreach ($file in $sourceFiles) {
    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    try {
        $component = $components.Item($componentName)

        if ($null -eq $component) {
            $failedComponents += $componentName
            Write-Host "  [FAIL] $componentName" -ForegroundColor Red
        }
        else {
            Write-Host "  [OK] $componentName" -ForegroundColor Green
        }
    }
    catch {
        $failedComponents += $componentName
        Write-Host "  [FAIL] $componentName" -ForegroundColor Red
    }
}

if ($failedComponents.Count -gt 0) {
    Write-Host ""
    Write-Host "Component gagal diverifikasi:" -ForegroundColor Red

    foreach ($name in $failedComponents) {
        Write-Host "  - $name" -ForegroundColor Red
    }

    exit 1
}

Write-Host ""
Write-Host "Verifikasi modDevTest..." -ForegroundColor Cyan

$devTest = $null

try {
    $devTest = $components.Item("modDevTest")
}
catch {
}

if ($null -eq $devTest) {
    Write-Host "  [FAIL] modDevTest tidak ditemukan." -ForegroundColor Red
    exit 1
}

Write-Host "  [OK] modDevTest ditemukan." -ForegroundColor Green

Write-Host ""
Write-Host "Menjalankan DevTest..." -ForegroundColor Cyan

try {
    $corel.GMSManager.RunMacro(
        "MgoCorel",
        "modDevTest.Plugin_DevTest"
    ) | Out-Null

    Write-Host "  [OK] Plugin_DevTest berhasil dijalankan." -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] Plugin_DevTest gagal dijalankan." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Menyimpan MgoCorel..." -ForegroundColor Cyan

$gmsProject.Dirty = $true

Start-Sleep -Milliseconds 500

try {
    $corel.VBE.ActiveVBProject
}
catch {
}

Start-Sleep -Milliseconds 500

$gmsPath = $gmsProject.FullFileName

if (-not (Test-Path $gmsPath)) {
    Write-Host ""
    Write-Host "[FAIL] MgoCorel.gms tidak ditemukan." -ForegroundColor Red
    exit 1
}

$fileInfo = Get-Item $gmsPath

Write-Host ""
Write-Host "Verifikasi GMS..." -ForegroundColor Cyan
Write-Host "  [OK] MgoCorel.gms ditemukan." -ForegroundColor Green
Write-Host "  Size: $($fileInfo.Length) bytes"
Write-Host "  Last Write: $($fileInfo.LastWriteTime)"

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "             BUILD OK" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Source berhasil di-import ke CorelDRAW."
Write-Host "Component berhasil diverifikasi."
Write-Host "modDevTest berhasil ditemukan."
Write-Host "MgoCorel.gms ditemukan."
Write-Host ""
#requires -Version 5.1
<#
    Сборка переносимого пакета dist\ (для установки на другом компьютере без прав администратора).
    Собирает решение, копирует DLL, скрипты, .bat и документацию, при необходимости архивирует в ZIP.
#>
param(
    [switch]$NoZip
)

$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root "dist"

Write-Host "== Сборка пакета Telegram Share ==" -ForegroundColor Cyan

# --- поиск MSBuild ---
$msbuild = $null
$vswhere = "C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe"
if ( Test-Path -Path $vswhere ) {
    $msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}
if ( -not $msbuild ) { $msbuild = "msbuild.exe" }

Write-Host ( "MSBuild: " + $msbuild )
& $msbuild ( Join-Path $root "TelegramShareAddin.sln" ) /restore /p:Configuration=Release /v:minimal /nologo
if ( $LASTEXITCODE -ne 0 ) { throw "Сборка завершилась с ошибкой." }

# --- сборка dist ---
if ( Test-Path -Path $dist ) { Remove-Item -Path $dist -Recurse -Force }
New-Item -ItemType Directory -Force -Path $dist | Out-Null
New-Item -ItemType Directory -Force -Path ( Join-Path $dist "scripts" ) | Out-Null

$wordBin = Join-Path $root "src\TelegramShareAddin\bin\Release"
$excelBin = Join-Path $root "src\TelegramShareExcelAddin\bin\Release"

Get-ChildItem -Path $wordBin -Filter "*.dll" | Copy-Item -Destination $dist -Force
Get-ChildItem -Path $excelBin -Filter "*.dll" | Copy-Item -Destination $dist -Force

foreach ( $name in @( "install.ps1", "uninstall.ps1", "check_install.ps1" ) ) {
    Copy-Item -Path ( Join-Path $root "scripts\$name" ) -Destination ( Join-Path $dist "scripts" ) -Force
}

foreach ( $name in @( "install.bat", "uninstall.bat", "check_install.bat" ) ) {
    if ( Test-Path -Path ( Join-Path $root $name ) ) {
        Copy-Item -Path ( Join-Path $root $name ) -Destination $dist -Force
    }
}

foreach ( $name in @( "README.md", "INSTALL.md", "LICENSE" ) ) {
    if ( Test-Path -Path ( Join-Path $root $name ) ) {
        Copy-Item -Path ( Join-Path $root $name ) -Destination $dist -Force
    }
}

$files = Get-ChildItem -Path $dist -File | Select-Object -ExpandProperty Name
Write-Host ( "В пакете dist: " + ( $files -join ", " ) ) -ForegroundColor Green

if ( -not $NoZip ) {
    $zip = Join-Path $root "TelegramShareAddin.zip"
    if ( Test-Path -Path $zip ) { Remove-Item -Path $zip -Force }
    Compress-Archive -Path ( Join-Path $dist "*" ) -DestinationPath $zip -Force
    Write-Host ( "Архив: " + $zip ) -ForegroundColor Green
}

Write-Host "Готово."
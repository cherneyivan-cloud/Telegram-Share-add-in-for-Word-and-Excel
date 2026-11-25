#requires -Version 5.1
<#
    Удаление надстроек Telegram Share (Word, Excel и Outlook).

    Программа ставится в %APPDATA%\TelegramShareAddin, данные пользователя
    (config.json, addin.log) лежат в %LOCALAPPDATA%\TelegramShareAddin и по умолчанию
    сохраняются: удалить их можно ключом -RemoveData.
#>
param(
    [switch]$RemoveFiles,
    [switch]$RemoveData
)

$ErrorActionPreference = "SilentlyContinue"

$InstallDir = Join-Path $env:APPDATA "TelegramShareAddin"
$DataDir = Join-Path $env:LOCALAPPDATA "TelegramShareAddin"

$items = @(
    @{ ProgId = "TelegramShareAddin"; Clsid = "{AD3B6C2F-4E5F-4A8B-9C1D-3E4F5A6B7C8D}"; App = "Word" } ,
    @{ ProgId = "TelegramShareExcel"; Clsid = "{B1C2D3E4-F5A6-4B7C-8D9E-0F1A2B3C4D5E}"; App = "Excel" } ,
    @{ ProgId = "TelegramShareOutlook"; Clsid = "{C2D3E4F5-A6B7-4C8D-9E0F-1A2B3C4D5E6F}"; App = "Outlook" }
)

Write-Host "== Удаление Telegram Share ==" -ForegroundColor Cyan

foreach ( $i in $items ) {
    Remove-Item -Path "HKCU:\Software\Microsoft\Office\$($i.App)\Addins\$($i.ProgId)" -Recurse -Force
    Remove-Item -Path "HKCU:\Software\Classes\CLSID\$($i.Clsid)" -Recurse -Force
    Remove-Item -Path "HKCU:\Software\Classes\$($i.ProgId)" -Recurse -Force
    Write-Host ( "  удалено: " + $i.App + " / " + $i.ProgId )
}

if ( Test-Path -Path $InstallDir ) {
    $doDelete = $RemoveFiles
    if ( -not $doDelete ) {
        $ans = Read-Host "Удалить файлы из $InstallDir? (y/N)"
        $doDelete = ( $ans -match "^[yY]" )
    }
    if ( $doDelete ) {
        Remove-Item -Path $InstallDir -Recurse -Force
        Write-Host "Файлы удалены."
    }
}

if ( $RemoveData ) {
    if ( Test-Path -Path $DataDir ) {
        Remove-Item -Path $DataDir -Recurse -Force
        Write-Host "Настройки и журнал удалены ($DataDir)."
    }
} elseif ( Test-Path -Path $DataDir ) {
    Write-Host "Настройки сохранены: $DataDir (удалить: uninstall.ps1 -RemoveData)." -ForegroundColor DarkGray
}

Write-Host "Готово. Перезапустите Word, Excel и Outlook."

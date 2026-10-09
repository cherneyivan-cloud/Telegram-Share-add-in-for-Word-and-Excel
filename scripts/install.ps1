#requires -Version 5.1
<#
    Установка надстроек Telegram Share для Microsoft Word и Microsoft Excel.
    Без прав администратора: всё пишется в HKCU и %APPDATA%.
    Поддерживает запуск из папки-пакета (dist) и из репозитория.
#>
param(
    [string]$SourcePath = "",
    [switch]$Yes
)

$ErrorActionPreference = "Stop"

$InstallDir = Join-Path $env:APPDATA "TelegramShareAddin"

function Find-Source( $dllName ) {
    $cands = @(
        $PSScriptRoot,
        ( Join-Path $PSScriptRoot ".." ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareAddin\bin\Release" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareAddin\bin\x64\Release" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareAddin\bin\Debug" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareAddin\bin\x64\Debug" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareExcelAddin\bin\Release" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareExcelAddin\bin\x64\Release" ),
        ( Join-Path $PSScriptRoot "..\src\TelegramShareExcelAddin\bin\Debug" )
    )
    foreach ( $c in $cands ) {
        if ( $c -and ( Test-Path -Path ( Join-Path $c $dllName ) ) ) {
            return ( Resolve-Path -Path $c ).Path
        }
    }
    return $null
}

function Register-Addin( $progId, $clsid, $className, $assemblyName, $dllPath, $officeApp, $friendlyName ) {
    $clsidKey = "HKCU:\Software\Classes\CLSID\$clsid"
    New-Item -Path $clsidKey -Force | Out-Null

    $inprocKey = "$clsidKey\InprocServer32"
    New-Item -Path $inprocKey -Force | Out-Null
    Set-ItemProperty -Path $inprocKey -Name "(default)" -Value "mscoree.dll"
    Set-ItemProperty -Path $inprocKey -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path $inprocKey -Name "Class" -Value $className
    Set-ItemProperty -Path $inprocKey -Name "Assembly" -Value ( $assemblyName + ", Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" )
    Set-ItemProperty -Path $inprocKey -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $inprocKey -Name "CodeBase" -Value ( "file:///" + ( $dllPath -replace "\\", "/" ) )

    $progKey = "HKCU:\Software\Classes\$progId"
    New-Item -Path $progKey -Force | Out-Null
    New-Item -Path "$progKey\CLSID" -Force | Out-Null
    Set-ItemProperty -Path "$progKey\CLSID" -Name "(default)" -Value $clsid

    $addinKey = "HKCU:\Software\Microsoft\Office\$officeApp\Addins\$progId"
    New-Item -Path $addinKey -Force | Out-Null
    Set-ItemProperty -Path $addinKey -Name "FriendlyName" -Value $friendlyName
    Set-ItemProperty -Path $addinKey -Name "Description" -Value "Отправка документа в Telegram Desktop"
    Set-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -Type DWord

    Write-Host ( "  [" + $officeApp + "] " + $progId + " -> " + $dllPath ) -ForegroundColor Green
}

Write-Host "== Установка Telegram Share — Word и Excel ==" -ForegroundColor Cyan

$wordSrc = Find-Source "TelegramShareAddin.dll"
$excelSrc = Find-Source "TelegramShareExcel.dll"

if ( -not $wordSrc -and -not $excelSrc ) {
    throw "Не найдены TelegramShareAddin.dll / TelegramShareExcel.dll. Укажите -SourcePath к папке со сборкой."
}

# --- Word/Excel должны быть закрыты ---
$busy = Get-Process -Name WINWORD,EXCEL -ErrorAction SilentlyContinue
if ( $busy -and ( -not $Yes ) ) {
    Write-Host "Обнаружены запущенные Word/Excel. Перезапустите их после установки." -ForegroundColor Yellow
    $ans = Read-Host "Продолжить установку? (y/N)"
    if ( $ans -notmatch "^[yY]" ) { Write-Host "Установка отменена."; exit 0 }
}

# --- Копирование файлов ---
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
foreach ( $src in @( $wordSrc, $excelSrc ) ) {
    if ( $src ) {
        Get-ChildItem -Path $src -Filter "*.dll" | Copy-Item -Destination $InstallDir -Force
    }
}
Write-Host "Файлы скопированы в $InstallDir" -ForegroundColor Green

# --- Регистрация COM + подключение к Office ---
if ( $wordSrc ) {
    Register-Addin "TelegramShareAddin" "{AD3B6C2F-4E5F-4A8B-9C1D-3E4F5A6B7C8D}" "TelegramShareAddin.Addin" `
        "TelegramShareAddin" ( Join-Path $InstallDir "TelegramShareAddin.dll" ) "Word" "Отправить в Telegram"
}
if ( $excelSrc ) {
    Register-Addin "TelegramShareExcel" "{B1C2D3E4-F5A6-4B7C-8D9E-0F1A2B3C4D5E}" "TelegramShareAddin.ExcelAddin" `
        "TelegramShareExcel" ( Join-Path $InstallDir "TelegramShareExcel.dll" ) "Excel" "Отправить в Telegram"
}

# --- Сброс списка отключённых надстроек Office ---
foreach ( $app in @( "Word", "Excel" ) ) {
    $resiliency = "HKCU:\Software\Microsoft\Office\16.0\$app\Resiliency"
    if ( Test-Path -Path $resiliency ) {
        Remove-Item -Path "$resiliency\DisabledItems" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$resiliency\CrashingAddinList" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- Ярлык «Отправить → Telegram» (SendTo) для любых файлов ---
$tgExe = Join-Path $env:APPDATA "Telegram Desktop\Telegram.exe"
if ( Test-Path -Path $tgExe ) {
    $sendto = Join-Path $env:APPDATA "Microsoft\Windows\SendTo"
    New-Item -ItemType Directory -Force -Path $sendto | Out-Null
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut( ( Join-Path $sendto "Telegram.lnk" ) )
    $lnk.TargetPath = $tgExe
    $lnk.Arguments = "-sendpath"
    $lnk.WorkingDirectory = ( Split-Path -Path $tgExe )
    $lnk.Save( )
    Write-Host "Ярлык «Отправить → Telegram» создан." -ForegroundColor Green
}

Write-Host ""
Write-Host "Готово. Перезапустите Word и/или Excel — вкладка «Отправить в Telegram» появится на ленте."
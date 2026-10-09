#requires -Version 5.1
<#
    Проверка установки надстроек Telegram Share (Word и Excel).
#>

$ErrorActionPreference = "Continue"

$InstallDir = Join-Path $env:APPDATA "TelegramShareAddin"
$allOk = $true

Write-Host "== Проверка установки Telegram Share ==" -ForegroundColor Cyan

$items = @(
    @{ ProgId = "TelegramShareAddin"; Clsid = "{AD3B6C2F-4E5F-4A8B-9C1D-3E4F5A6B7C8D}"; App = "Word"; Dll = "TelegramShareAddin.dll" } ,
    @{ ProgId = "TelegramShareExcel"; Clsid = "{B1C2D3E4-F5A6-4B7C-8D9E-0F1A2B3C4D5E}"; App = "Excel"; Dll = "TelegramShareExcel.dll" }
)

foreach ( $i in $items ) {
    Write-Host ( "--- " + $i.App + " (" + $i.ProgId + ") ---" )

    $dllPath = Join-Path $InstallDir $i.Dll
    if ( Test-Path -Path $dllPath ) {
        Write-Host ( "  [OK] файл: " + $dllPath ) -ForegroundColor Green
    }
    else {
        Write-Host ( "  [FAIL] нет файла " + $dllPath ) -ForegroundColor Red
        $allOk = $false
    }

    $inproc = "HKCU:\Software\Classes\CLSID\$($i.Clsid)\InprocServer32"
    if ( Test-Path -Path $inproc ) {
        $server = ( Get-Item -Path $inproc ).GetValue( "" )
        $cls = ( Get-ItemProperty -Path $inproc -Name "Class" -ErrorAction SilentlyContinue ).Class
        if ( $server -eq "mscoree.dll" ) { Write-Host "  [OK] InprocServer32 = mscoree.dll" -ForegroundColor Green }
        else { Write-Host ( "  [WARN] InprocServer32 = " + $server ) -ForegroundColor Yellow }
        Write-Host ( "  [OK] Class = " + $cls ) -ForegroundColor Green
    }
    else {
        Write-Host ( "  [FAIL] нет ключа " + $inproc ) -ForegroundColor Red
        $allOk = $false
    }

    $addinKey = "HKCU:\Software\Microsoft\Office\$($i.App)\Addins\$($i.ProgId)"
    if ( Test-Path -Path $addinKey ) {
        $lb = ( Get-ItemProperty -Path $addinKey ).LoadBehavior
        if ( $lb -eq 3 ) { Write-Host "  [OK] LoadBehavior = 3" -ForegroundColor Green }
        else { Write-Host ( "  [WARN] LoadBehavior = " + $lb ) -ForegroundColor Yellow }
    }
    else {
        Write-Host ( "  [FAIL] нет ключа " + $addinKey ) -ForegroundColor Red
        $allOk = $false
    }

    $obj = $null
    try {
        $obj = New-Object -ComObject $i.ProgId
        Write-Host "  [OK] COM-объект создаётся" -ForegroundColor Green
    }
    catch {
        Write-Host ( "  [FAIL] COM-объект не создаётся: " + $_.Exception.Message ) -ForegroundColor Red
        $allOk = $false
    }
    finally {
        if ( $null -ne $obj ) {
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject( $obj ) } catch { }
        }
    }
}

Write-Host ""
if ( $allOk ) {
    Write-Host "ИТОГ: установка корректна. Перезапустите Word и Excel." -ForegroundColor Green
}
else {
    Write-Host "ИТОГ: обнаружены проблемы. Запустите install.ps1 и перезапустите Office." -ForegroundColor Red
}
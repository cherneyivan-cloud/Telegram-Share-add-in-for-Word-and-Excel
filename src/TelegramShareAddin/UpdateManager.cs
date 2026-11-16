using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace TelegramShareAddin
{
    /// <summary>Проверка и установка обновлений из GitHub Releases.</summary>
    internal static class UpdateManager
    {
        public const string CurrentVersion = "1.5.3";

        private const string LatestReleaseUrl =
            "https://api.github.com/repos/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases/latest";

        public static void CheckAndInstall( )
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                ReleaseInfo release = GetLatestRelease( );
                Cursor.Current = Cursors.Default;

                if ( release == null || release.Version == null )
                {
                    MessageBox.Show(
                        "Не удалось определить последнюю версию.",
                        "Telegram Share",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning );
                    return;
                }

                Version current = new Version( CurrentVersion );
                if ( release.Version <= current )
                {
                    MessageBox.Show(
                        "Установлена последняя версия: v" + CurrentVersion + ".",
                        "Telegram Share",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information );
                    return;
                }

                DialogResult answer = MessageBox.Show(
                    "Доступно обновление v" + release.Version + ".\r\n" +
                    "Установить его после закрытия Word, Excel и Outlook?",
                    "Telegram Share — обновление",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information );

                if ( answer != DialogResult.Yes )
                {
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;
                string updateRoot = Path.Combine(
                    Path.GetTempPath( ),
                    "TelegramShareUpdate_" + Guid.NewGuid( ).ToString( "N" ) );
                Directory.CreateDirectory( updateRoot );

                string zipPath = Path.Combine( updateRoot , "TelegramShareAddin.zip" );
                using ( var client = CreateClient( ) )
                {
                    client.DownloadFile( release.AssetUrl , zipPath );
                }

                string scriptPath = Path.Combine( updateRoot , "install-update.ps1" );
                File.WriteAllText( scriptPath , BuildInstallerScript( ) , new UTF8Encoding( true ) );

                string args =
                    "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal " +
                    "-File " + Quote( scriptPath ) +
                    " -Zip " + Quote( zipPath ) +
                    " -Work " + Quote( updateRoot );

                Process.Start( new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = false
                } );

                Cursor.Current = Cursors.Default;
                MessageBox.Show(
                    "Обновление скачано.\r\n\r\n" +
                    "Закройте Word, Excel и Outlook. Установщик автоматически заменит файлы " +
                    "и завершит обновление.",
                    "Telegram Share",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                Cursor.Current = Cursors.Default;
                AddinCore.Log( "Update ERROR: " + ex.ToString( ) );
                MessageBox.Show(
                    "Не удалось проверить или скачать обновление:\r\n" + ex.Message,
                    "Telegram Share",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error );
            }
        }

        private static ReleaseInfo GetLatestRelease( )
        {
            using ( var client = CreateClient( ) )
            {
                string json = client.DownloadString( LatestReleaseUrl );
                JObject root = JObject.Parse( json );
                string tag = (string)root[ "tag_name" ];
                Version version;
                if ( string.IsNullOrEmpty( tag ) ||
                     !Version.TryParse( tag.Trim( ).TrimStart( 'v' , 'V' ) , out version ) )
                {
                    return null;
                }

                string assetUrl = null;
                JArray assets = root[ "assets" ] as JArray;
                if ( assets != null )
                {
                    foreach ( JObject asset in assets )
                    {
                        if ( string.Equals( (string)asset[ "name" ] , "TelegramShareAddin.zip" , StringComparison.OrdinalIgnoreCase ) )
                        {
                            assetUrl = (string)asset[ "browser_download_url" ];
                            break;
                        }
                    }
                }

                if ( string.IsNullOrEmpty( assetUrl ) )
                {
                    return null;
                }

                return new ReleaseInfo { Version = version , AssetUrl = assetUrl };
            }
        }

        private static WebClient CreateClient( )
        {
            var client = new WebClient( );
            client.Headers[ HttpRequestHeader.UserAgent ] = "TelegramShareAddin/" + CurrentVersion;
            return client;
        }

        private static string BuildInstallerScript( )
        {
            return @"
param(
    [Parameter(Mandatory = $true)][string]$Zip,
    [Parameter(Mandatory = $true)][string]$Work
)

$ErrorActionPreference = 'Continue'
$log = Join-Path $Work 'update.log'
Start-Transcript -Path $log -Append | Out-Null

Write-Host 'Telegram Share: ожидаю закрытия Office...'
while ( Get-Process -Name WINWORD,EXCEL,OUTLOOK -ErrorAction SilentlyContinue ) {
    Start-Sleep -Seconds 2
}

$extract = Join-Path $Work 'package'
Expand-Archive -Path $Zip -DestinationPath $extract -Force
$install = Join-Path $extract 'scripts\install.ps1'

if ( Test-Path $install ) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $install -Yes
    Write-Host 'Обновление завершено. Можно запускать Office.'
} else {
    Write-Host 'Не найден scripts\install.ps1 в архиве.'
}

Stop-Transcript | Out-Null
Start-Sleep -Seconds 3
Remove-Item -Path $Work -Recurse -Force -ErrorAction SilentlyContinue
";
        }

        private static string Quote( string value )
        {
            return "\"" + value.Replace( "\"" , "\\\"" ) + "\"";
        }

        private sealed class ReleaseInfo
        {
            public Version Version { get; set; }
            public string AssetUrl { get; set; }
        }
    }
}

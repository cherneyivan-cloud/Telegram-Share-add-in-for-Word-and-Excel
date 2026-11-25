using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace TelegramShareAddin
{
    /// <summary>Проверка и установка обновлений из GitHub Releases.</summary>
    internal static class UpdateManager
    {
        /// <summary>Версия надстройки — берётся из сборки, чтобы не расходиться с AssemblyInfo.</summary>
        public static string CurrentVersion
        {
            get
            {
                try
                {
                    Version version = typeof( UpdateManager ).Assembly.GetName( ).Version;
                    return version.Major + "." + version.Minor + "." + version.Build;
                }
                catch
                {
                    return "0.0.0";
                }
            }
        }

        private const string LatestReleaseUrl =
            "https://api.github.com/repos/cherneyivan-cloud/Telegram-Share-add-in-for-Word-and-Excel/releases/latest";

        private const string AssetName = "TelegramShareAddin.zip";

        private static HttpClient _http;

        private static HttpClient Http
        {
            get
            {
                if ( _http == null )
                {
                    EnableModernTls( );

                    var client = new HttpClient( );
                    client.Timeout = TimeSpan.FromSeconds( 45 );
                    client.DefaultRequestHeaders.UserAgent.ParseAdd( "TelegramShareAddin/" + CurrentVersion );
                    client.DefaultRequestHeaders.Accept.ParseAdd( "application/vnd.github+json" );
                    _http = client;
                }

                return _http;
            }
        }

        /// <summary>GitHub отклоняет TLS 1.0/1.1; TLS 1.1 в список не добавляем.</summary>
        private static void EnableModernTls( )
        {
            try
            {
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
            }
            catch
            {
                try
                {
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                }
                catch
                {
                }
            }
        }

        public static async void CheckAndInstall( )
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                ReleaseInfo release = await GetLatestReleaseAsync( );
                Cursor.Current = Cursors.Default;

                if ( release == null || release.Version == null )
                {
                    MessageBox.Show(
                        "Не удалось определить последнюю версию." ,
                        "Telegram Share" ,
                        MessageBoxButtons.OK ,
                        MessageBoxIcon.Warning );
                    return;
                }

                var current = new Version( CurrentVersion );
                if ( release.Version <= current )
                {
                    MessageBox.Show(
                        "Установлена последняя версия: v" + CurrentVersion + "." ,
                        "Telegram Share" ,
                        MessageBoxButtons.OK ,
                        MessageBoxIcon.Information );
                    return;
                }

                DialogResult answer = MessageBox.Show(
                    "Доступно обновление v" + release.Version + ".\r\n" +
                    "Установить его после закрытия Word, Excel и Outlook?" ,
                    "Telegram Share — обновление" ,
                    MessageBoxButtons.YesNo ,
                    MessageBoxIcon.Information );

                if ( answer != DialogResult.Yes )
                {
                    return;
                }

                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    string updateRoot = Path.Combine(
                        Path.GetTempPath( ) ,
                        "TelegramShareUpdate_" + Guid.NewGuid( ).ToString( "N" ) );
                    Directory.CreateDirectory( updateRoot );

                    string zipPath = Path.Combine( updateRoot , AssetName );
                    byte[] package = await Http.GetByteArrayAsync( release.AssetUrl );
                    VerifyDigest( release.Digest , package );
                    File.WriteAllBytes( zipPath , package );

                    string scriptPath = Path.Combine( updateRoot , "install-update.ps1" );
                    File.WriteAllText( scriptPath , BuildInstallerScript( ) , new UTF8Encoding( true ) );

                    string args =
                        "-NoProfile -ExecutionPolicy Bypass -WindowStyle Normal " +
                        "-File " + Quote( scriptPath ) +
                        " -Zip " + Quote( zipPath ) +
                        " -Work " + Quote( updateRoot );

                    Process.Start( new ProcessStartInfo
                    {
                        FileName = "powershell.exe" ,
                        Arguments = args ,
                        UseShellExecute = false ,
                        CreateNoWindow = false
                    } );
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }

                MessageBox.Show(
                    "Обновление скачано и проверено.\r\n\r\n" +
                    "Закройте Word, Excel и Outlook. Установщик автоматически заменит файлы " +
                    "и завершит обновление." ,
                    "Telegram Share" ,
                    MessageBoxButtons.OK ,
                    MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                Cursor.Current = Cursors.Default;
                AddinLog.Error( "Update" , ex );
                MessageBox.Show(
                    "Не удалось проверить или скачать обновление:\r\n" + ex.Message ,
                    "Telegram Share" ,
                    MessageBoxButtons.OK ,
                    MessageBoxIcon.Error );
            }
        }

        /// <summary>Сверяет SHA-256 архива с дайджестом из GitHub Releases.</summary>
        private static void VerifyDigest( string digest , byte[] package )
        {
            if ( string.IsNullOrEmpty( digest ) )
            {
                // Без контрольной суммы установку не начинаем.
                throw new InvalidOperationException( "Релиз не содержит контрольной суммы — установка отменена." );
            }

            const string prefix = "sha256:";
            if ( !digest.StartsWith( prefix , StringComparison.OrdinalIgnoreCase ) )
            {
                throw new InvalidOperationException( "Неизвестный формат дайджеста обновления: " + digest );
            }

            string expected = digest.Substring( prefix.Length ).Trim( );

            using ( var sha = SHA256.Create( ) )
            {
                string actual = BitConverter.ToString( sha.ComputeHash( package ) ).Replace( "-" , "" );
                if ( !string.Equals( actual , expected , StringComparison.OrdinalIgnoreCase ) )
                {
                    throw new InvalidOperationException(
                        "Контрольная сумма обновления не совпала. Загрузка отменена." );
                }
            }
        }

        private static async Task<ReleaseInfo> GetLatestReleaseAsync( )
        {
            string json = await Http.GetStringAsync( LatestReleaseUrl );
            JObject root = JObject.Parse( json );
            string tag = (string) root[ "tag_name" ];

            Version version;
            if ( string.IsNullOrEmpty( tag ) ||
                 !Version.TryParse( tag.Trim( ).TrimStart( 'v' , 'V' ) , out version ) )
            {
                return null;
            }

            string assetUrl = null;
            string digest = null;

            JArray assets = root[ "assets" ] as JArray;
            if ( assets != null )
            {
                foreach ( JObject asset in assets )
                {
                    if ( string.Equals( (string) asset[ "name" ] , AssetName , StringComparison.OrdinalIgnoreCase ) )
                    {
                        assetUrl = (string) asset[ "browser_download_url" ];
                        digest = (string) asset[ "digest" ];
                        break;
                    }
                }
            }

            if ( string.IsNullOrEmpty( assetUrl ) )
            {
                return null;
            }

            return new ReleaseInfo { Version = version , AssetUrl = assetUrl , Digest = digest };
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
            // В PowerShell безопаснее одинарные кавычки: внутри них кавычка удваивается.
            return "'" + ( value ?? string.Empty ).Replace( "'" , "''" ) + "'";
        }

        private sealed class ReleaseInfo
        {
            public Version Version { get; set; }
            public string AssetUrl { get; set; }
            public string Digest { get; set; }
        }
    }
}

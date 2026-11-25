using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;

namespace TelegramShareAddin
{
    /// <summary>
    /// Настройки надстройки.
    ///
    /// Лежат в %LOCALAPPDATA%\TelegramShareAddin\config.json — отдельно от папки установки,
    /// чтобы удаление или обновление надстройки не уносило контакты. Старый файл из
    /// %APPDATA%\TelegramShareAddin переносится автоматически при первом запуске.
    /// </summary>
    public class Settings
    {
        public const string AppName = "TelegramShareAddin";

        public static string ConfigDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData ) ,
                    AppName );
            }
        }

        public static string ConfigFilePath
        {
            get
            {
                return Path.Combine( ConfigDirectory , "config.json" );
            }
        }

        /// <summary>Прежнее расположение настроек — только для переноса.</summary>
        private static string LegacyConfigFilePath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) ,
                    AppName ,
                    "config.json" );
            }
        }

        /// <summary>Последние использованные контакты (для автодополнения в диалоге).</summary>
        public List<string> RecentContacts { get; set; } = new List<string>( );

        /// <summary>Чаты, скрытые пользователем в списке (чтобы не искать их повторно).</summary>
        public List<string> HiddenChats { get; set; } = new List<string>( );

        /// <summary>Последний выбранный формат: "PDF" или "DOCX".</summary>
        public string LastFormat { get; set; } = "PDF";


        /// <summary>Задержка после открытия чата перед вставкой файла, мс.</summary>
        public int ChatOpenDelayMs { get; set; } = 2500;

        /// <summary>Пауза после вставки файла перед набором сообщения/отправкой, мс.</summary>
        public int PasteDelayMs { get; set; } = 1200;

        /// <summary>Запоминать ли последний контакт и формат.</summary>
        public bool RememberContact { get; set; } = true;

        /// <summary>Задержка открытия чата с защитой от неверных значений в файле.</summary>
        public int ChatOpenDelay
        {
            get
            {
                return Clamp( ChatOpenDelayMs , 200 , 60000 , 2500 );
            }
        }

        /// <summary>Пауза после вставки с защитой от неверных значений в файле.</summary>
        public int PasteDelay
        {
            get
            {
                return Clamp( PasteDelayMs , 200 , 60000 , 1200 );
            }
        }


        /// <summary>Загрузить настройки. Ошибки попадают в лог; битый файл сохраняется как .broken.</summary>
        public static Settings Load( )
        {
            // 1. Рабочий файл, 2. резервная копия рядом с ним.
            Settings loaded = ReadFile( ConfigFilePath ) ?? ReadFile( ConfigFilePath + ".bak" );
            if ( loaded != null )
            {
                return loaded;
            }

            // 3. Прежнее расположение — только если нового файла ещё нет (первый запуск).
            if ( !File.Exists( ConfigFilePath ) )
            {
                loaded = ReadFile( LegacyConfigFilePath );
                if ( loaded != null )
                {
                    AddinLog.Info( "Настройки перенесены из " + LegacyConfigFilePath );
                    loaded.Save( );
                    return loaded;
                }
            }

            return new Settings( );
        }

        private static Settings ReadFile( string path )
        {
            try
            {
                if ( !File.Exists( path ) )
                {
                    return null;
                }

                var settings = JsonConvert.DeserializeObject<Settings>( File.ReadAllText( path ) );
                if ( settings == null )
                {
                    return null;
                }

                settings.Normalize( );
                return settings;
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "Не удалось загрузить настройки (" + path + ")" , ex );

                // Битый файл не выбрасываем: пользователь сможет достать из него контакты.
                TryBackupBroken( path );
                return null;
            }
        }

        private static void TryBackupBroken( string path )
        {
            try
            {
                if ( File.Exists( path ) )
                {
                    File.Copy( path , path + ".broken" , true );
                    AddinLog.Warn( "Повреждённый файл настроек сохранён как " + path + ".broken" );
                }
            }
            catch
            {
            }
        }

        /// <summary>Сохранить настройки: атомарно и с резервной копией.</summary>
        public bool Save( )
        {
            try
            {
                Directory.CreateDirectory( ConfigDirectory );

                string json = JsonConvert.SerializeObject( this , Formatting.Indented );
                string temp = ConfigFilePath + ".tmp";

                WriteWithRetry( temp , json );

                if ( File.Exists( ConfigFilePath ) )
                {
                    File.Replace( temp , ConfigFilePath , ConfigFilePath + ".bak" , true );
                }
                else
                {
                    File.Move( temp , ConfigFilePath );
                }

                return true;
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "Не удалось сохранить настройки" , ex );
                return false;
            }
        }

        private static void WriteWithRetry( string path , string content )
        {
            Exception last = null;

            for ( int i = 0 ; i < 5 ; i++ )
            {
                try
                {
                    File.WriteAllText( path , content );
                    return;
                }
                catch ( IOException ex )
                {
                    last = ex;
                    Thread.Sleep( 120 );
                }
            }

            throw last ?? new IOException( "Не удалось записать " + path );
        }

        /// <summary>Привести значения к допустимым (файл мог быть отредактирован вручную).</summary>
        private void Normalize( )
        {
            ChatOpenDelayMs = Clamp( ChatOpenDelayMs , 200 , 60000 , 2500 );
            PasteDelayMs = Clamp( PasteDelayMs , 200 , 60000 , 1200 );

            if ( RecentContacts == null )
            {
                RecentContacts = new List<string>( );
            }

            if ( HiddenChats == null )
            {
                HiddenChats = new List<string>( );
            }

            if ( !string.Equals( LastFormat , "DOCX" , StringComparison.OrdinalIgnoreCase ) )
            {
                LastFormat = "PDF";
            }
            else
            {
                LastFormat = "DOCX";
            }


        }

        private static int Clamp( int value , int min , int max , int fallback )
        {
            if ( value <= 0 )
            {
                return fallback;
            }

            if ( value < min )
            {
                return min;
            }

            return value > max ? max : value;
        }

        /// <summary>Добавить контакт в историю (максимум 20, без повторов).</summary>
        public void RememberContactName( string contact )
        {
            if ( contact == null )
            {
                return;
            }

            string c = contact.Trim( ).TrimStart( '@' );
            if ( string.IsNullOrEmpty( c ) )
            {
                return;
            }

            RecentContacts.RemoveAll( x => string.Equals( x , c , StringComparison.OrdinalIgnoreCase ) );
            RecentContacts.Insert( 0 , c );
            if ( RecentContacts.Count > 20 )
            {
                RecentContacts.RemoveRange( 20 , RecentContacts.Count - 20 );
            }
        }

        /// <summary>Скрыть чат из списка выбора.</summary>
        public void HideChat( string chat )
        {
            if ( string.IsNullOrEmpty( chat ) )
            {
                return;
            }

            if ( !HiddenChats.Exists( x => string.Equals( x , chat , StringComparison.OrdinalIgnoreCase ) ) )
            {
                HiddenChats.Add( chat );
            }
        }
    }
}

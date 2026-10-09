using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace TelegramShareAddin
{
    /// <summary>Настройки надстройки. Хранятся в %APPDATA%\TelegramShareAddin\config.json.</summary>
    public class Settings
    {
        public const string AppName = "TelegramShareAddin";

        public static string ConfigDirectory
        {
            get
            {
                string baseDir = Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData );
                return Path.Combine( baseDir, AppName );
            }
        }

        public static string ConfigFilePath
        {
            get
            {
                return Path.Combine( ConfigDirectory, "config.json" );
            }
        }

        /// <summary>Последние использованные контакты (для автодополнения в диалоге(.</summary>
        public List<string> RecentContacts { get; set; } = new List<string>( );

        /// <summary>Последний выбранный формат: "PDF" или "DOCX".</summary>
        public string LastFormat { get; set; } = "PDF";

        /// <summary>Задержка после открытия чата Telegram перед вставкой файла, мс.</summary>
        public int ChatOpenDelayMs { get; set; } = 2500;

        /// <summary>Пауза после вставки файла перед набором сообщения/отправкой, мс.</summary>
        public int PasteDelayMs { get; set; } = 1200;

        /// <summary>Запоминать ли последний контакт и формат.</summary>
        public bool RememberContact { get; set; } = true;

        /// <summary>Загрузить настройки. При ошибке — вернуть настройки по умолчанию.</summary>
        public static Settings Load( )
        {
            try
            {
                if ( File.Exists( ConfigFilePath ) )
                {
                    var s = JsonConvert.DeserializeObject<Settings>( File.ReadAllText( ConfigFilePath ) );
                    if ( s != null )
                    {
                        return s;
                    }
                }
            }
            catch ( Exception ex )
            {
                MessageBox.Show( "Не удалось загрузить настройки: " + ex.Message +
                    "\r\nБудут использованы настройки по умолчанию." ,
                    "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Warning );
            }

            return new Settings( );
        }

        /// <summary>Сохранить настройки.</summary>
        public void Save( )
        {
            try
            {
                Directory.CreateDirectory( ConfigDirectory );
                File.WriteAllText( ConfigFilePath , JsonConvert.SerializeObject( this , Formatting.Indented ) );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( "Не удалось сохранить настройки: " + ex.Message ,
                    "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Warning );
            }
        }

        /// <summary>Добавить контакт в историю (максимум 20{, без повторов(.</summary>
        public void RememberContactName( string contact )
        {
            string c = contact.Trim( ).TrimStart( '@' );
            if ( string.IsNullOrEmpty( c ) )
            {
                return;
            }

            RecentContacts.RemoveAll( x => string.Equals( x,c , StringComparison.OrdinalIgnoreCase ) );
            RecentContacts.Insert( 0,c );
            if ( RecentContacts.Count >   20 )
            {
                RecentContacts.RemoveRange( 20 ,  RecentContacts.Count -20 );
            }
        }
    }
}
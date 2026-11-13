using System;
using System.Collections.Specialized;
using System.Threading;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Работа с буфером обмена.
    ///
    /// Файл кладётся вместе с полным путём текстом: в Telegram/MAX Ctrl+V вставит
    /// файл, а в любом текстовом поле — полный путь к нему.
    /// </summary>
    internal static class ClipboardHelper
    {
        private const int RetryCount = 10;
        private const int RetryDelayMs = 100;

        /// <summary>Положить файл в буфер и продублировать его полный путь текстом.</summary>
        public static void SetFileAndPath( string filePath )
        {
            var files = new StringCollection( );
            files.Add( filePath );

            var data = new DataObject( );
            data.SetFileDropList( files );
            data.SetText( filePath );

            SetDataObject( data );
        }

        /// <summary>Положить в буфер только текст (например, полный путь к файлу).</summary>
        public static void SetText( string text )
        {
            SetDataObject( text ?? string.Empty );
        }

        private static void SetDataObject( object data )
        {
            Exception last = null;

            for ( int i = 0; i < RetryCount; i++ )
            {
                try
                {
                    Clipboard.SetDataObject( data , true );
                    return;
                }
                catch ( Exception ex )
                {
                    last = ex;
                    Thread.Sleep( RetryDelayMs );
                }
            }

            throw new InvalidOperationException(
                "Не удалось поместить данные в буфер обмена: " + ( last != null ? last.Message : "неизвестная ошибка" ) ,
                last );
        }
    }
}

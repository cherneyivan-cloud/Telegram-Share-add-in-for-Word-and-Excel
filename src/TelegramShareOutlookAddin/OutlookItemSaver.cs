using System;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TelegramShareAddin
{
    /// <summary>
    /// Сохранение письма Outlook во временный .msg.
    /// PDF из Outlook штатно не экспортируется, поэтому формат всегда MSG.
    /// </summary>
    internal static class OutlookItemSaver
    {
        public static string Save( object item , SendFileFormat format )
        {
            Outlook.MailItem mail = item as Outlook.MailItem;
            if ( mail == null )
            {
                throw new InvalidOperationException( "Поддерживаются только письма Outlook. Выделите письмо и повторите." );
            }

            string subject = null;
            try
            {
                subject = mail.Subject;
            }
            catch
            {
            }

            string baseName = Sanitize( subject );
            if ( string.IsNullOrEmpty( baseName ) )
            {
                baseName = "Письмо";
            }

            string path = TelegramSender.NewTempFile( baseName + ".msg" );
            TelegramSender.CleanupOldTempFiles( );

            try
            {
                mail.SaveAs( path , Outlook.OlSaveAsType.olMSG );
            }
            catch ( Exception ex )
            {
                throw new InvalidOperationException( "Не удалось сохранить письмо в MSG: " + ex.Message , ex );
            }

            try
            {
                if ( !File.Exists( path ) || ( new FileInfo( path ) ).Length == 0 )
                {
                    throw new InvalidOperationException( "Временный файл не создан или пуст: " + path );
                }
            }
            catch ( InvalidOperationException )
            {
                throw;
            }
            catch
            {
            }

            return path;
        }

        private static string Sanitize( string name )
        {
            if ( string.IsNullOrEmpty( name ) )
            {
                return null;
            }

            foreach ( char c in Path.GetInvalidFileNameChars( ) )
            {
                name = name.Replace( c , '_' );
            }

            name = name.Trim( );
            if ( name.Length > 120 )
            {
                name = name.Substring( 0 , 120 );
            }
            return name;
        }
    }
}
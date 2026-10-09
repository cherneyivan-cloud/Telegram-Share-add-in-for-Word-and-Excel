using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace TelegramShareAddin
{
    /// <summary>
    /// Общая (статическая) логика надстроек Word и Excel.
    /// COM-интерфейсы и методы ленты реализуются в самих классах надстроек.
    /// </summary>
    internal static class AddinCore
    {
        private static IPictureDisp _imageRectangular;
        private static IPictureDisp _imageRounded;
        private static IPictureDisp _imageGitHub;

        public static string GetCustomUi( string ribbonId )
        {
            try
            {
                Log( "GetCustomUI: " + ribbonId );
                if ( string.IsNullOrEmpty( ribbonId ) || ribbonId.StartsWith( "Microsoft." , StringComparison.OrdinalIgnoreCase ) )
                {
                    return RibbonDefinition.Xml;
                }
                return null;
            }
            catch ( Exception ex )
            {
                Log( "GetCustomUI ERROR: " + ex.ToString( ) );
                return null;
            }
        }

        public static object GetImage( IRibbonControl control )
        {
            try
            {
                if ( control != null && control.Id == "TelegramShareSendViaButton" )
                {
                    if ( _imageRounded == null )
                    {
                        _imageRounded = PictureConverter.ToPictureDisp( IconFactory.CreateRounded( ) );
                    }
                    return _imageRounded;
                }

                if ( control != null && control.Id == "TelegramShareButton" )
                {
                    if ( _imageRectangular == null )
                    {
                        _imageRectangular = PictureConverter.ToPictureDisp( IconFactory.CreateRectangular( ) );
                    }
                    return _imageRectangular;
                }

                if ( control != null && control.Id == "TelegramShareGitHubButton" )
                {
                    if ( _imageGitHub == null )
                    {
                        _imageGitHub = PictureConverter.ToPictureDisp( IconFactory.CreateGitHub( ) );
                    }
                    return _imageGitHub;
                }
            }
            catch ( Exception ex )
            {
                Log( "GetButtonImage ERROR: " + ex.ToString( ) );
            }
            return null;
        }

        public static void SendMessage( Settings settings , Func<bool> hasDocument , Func<SendFileFormat,string> saveDocument , string noDocumentMessage , string nativeLabel )
        {
            try
            {
                Log( "OnSendButton: enter" );

                if ( !hasDocument( ) )
                {
                    MessageBox.Show( noDocumentMessage , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                    return;
                }

                using ( var form = new SendForm( settings , nativeLabel ) )
                {
                    if ( form.ShowDialog( ) != DialogResult.OK )
                    {
                        return;
                    }

                    settings.RememberContact = form.RememberContact;
                    if ( settings.RememberContact )
                    {
                        settings.RememberContactName( form.Contact );
                    }
                    settings.LastFormat = form.Format == SendFileFormat.Docx ? "DOCX" : "PDF";
                    settings.Save( );

                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        string tempFile = saveDocument( form.Format );
                        var options = new TelegramSendOptions
                        {
                            Contact = form.Contact ,
                            Message = form.Message ,
                            Format = form.Format ,
                            OpenByChatName = form.OpenByChatName
                        };
                        var sender = new TelegramSender( settings );
                        sender.SendPreparedFile( options , tempFile );
                    }
                    finally
                    {
                        Cursor.Current = Cursors.Default;
                    }
                }

                MessageBox.Show( "Документ отправлен в Telegram." , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                Log( "OnSendButton ERROR: " + ex.ToString( ) );
                MessageBox.Show( "Ошибка: " + ex.Message , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
        }

        public static void SendVia( Settings settings , Func<bool> hasDocument , Func<SendFileFormat,string> saveDocument , string noDocumentMessage , string nativeLabel )
        {
            try
            {
                Log( "OnSendViaTelegram: enter" );

                if ( !hasDocument( ) )
                {
                    MessageBox.Show( noDocumentMessage , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                    return;
                }

                SendFileFormat format = settings.LastFormat == "DOCX" ? SendFileFormat.Docx : SendFileFormat.Pdf;

                using ( var formatForm = new FormatForm( format , nativeLabel ) )
                {
                    if ( formatForm.ShowDialog( ) != DialogResult.OK )
                    {
                        return;
                    }
                    format = formatForm.Format;
                }

                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    string tempFile = saveDocument( format );
                    var sender = new TelegramSender( settings );
                    sender.SendWithPickerPreparedFile( tempFile );
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }

                MessageBox.Show( "Документ передан в Telegram. Выберите чат в окне «Переслать…» и нажмите «Отправить»." , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                Log( "OnSendViaTelegram ERROR: " + ex.ToString( ) );
                MessageBox.Show( "Ошибка: " + ex.Message , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
        }

        public static void OpenGitHub( )
        {
            try
            {
                Log( "OnOpenGitHub: " + RibbonDefinition.GitHubUrl );
                System.Diagnostics.Process.Start( new System.Diagnostics.ProcessStartInfo( RibbonDefinition.GitHubUrl ) { UseShellExecute = true } );
            }
            catch ( Exception ex )
            {
                Log( "OnOpenGitHub ERROR: " + ex.ToString( ) );
                MessageBox.Show( "Не удалось открыть ссылку: " + ex.Message , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
        }

        public static void Log( string message )        {
            try
            {
                string dir = Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) , "TelegramShareAddin" );
                Directory.CreateDirectory( dir );
                string line = DateTime.Now.ToString( "yyyy-MM-dd HH:mm:ss.fff" ) + "  " + message + Environment.NewLine;
                File.AppendAllText( Path.Combine( dir , "addin.log" ) , line );
            }
            catch
            {
            }
        }
    }
}
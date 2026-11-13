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
        private static IPictureDisp _imagePaperclip;
        private static IPictureDisp _imagePrinter;
        private static IPictureDisp _imageMax;

        public static string GetCustomUi( string ribbonId , bool outlook )
        {
            try
            {
                Log( "GetCustomUI: " + ribbonId );
                if ( string.IsNullOrEmpty( ribbonId ) || ribbonId.StartsWith( "Microsoft." , StringComparison.OrdinalIgnoreCase ) )
                {
                    return outlook ? RibbonDefinition.OutlookXml : RibbonDefinition.Xml;
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

                if ( control != null && control.Id == "TelegramShareAttachmentsButton" )
                {
                    if ( _imagePaperclip == null )
                    {
                        _imagePaperclip = PictureConverter.ToPictureDisp( IconFactory.CreatePaperclip( ) );
                    }
                    return _imagePaperclip;
                }

                if ( control != null && control.Id == "TelegramSharePrintButton" )
                {
                    if ( _imagePrinter == null )
                    {
                        _imagePrinter = PictureConverter.ToPictureDisp( IconFactory.CreatePrinter( ) );
                    }
                    return _imagePrinter;
                }

                if ( control != null && control.Id == "TelegramShareMaxButton" )
                {
                    if ( _imageMax == null )
                    {
                        _imageMax = PictureConverter.ToPictureDisp( IconFactory.CreateMax( ) );
                    }
                    return _imageMax;
                }
            }
            catch ( Exception ex )
            {
                Log( "GetButtonImage ERROR: " + ex.ToString( ) );
            }
            return null;
        }

        public static void SendMessage( Settings settings , Func<bool> hasDocument , Func<SendFileFormat,string> saveDocument , string noDocumentMessage , string nativeLabel , bool allowPdf , Messenger defaultMessenger )
        {
            try
            {
                Log( "OnSendButton: enter, messenger = " + defaultMessenger );

                if ( !hasDocument( ) )
                {
                    MessageBox.Show( noDocumentMessage , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                    return;
                }

                Messenger messenger;

                using ( var form = new SendForm( settings , nativeLabel , allowPdf , defaultMessenger ) )
                {
                    if ( form.ShowDialog( ) != DialogResult.OK )
                    {
                        return;
                    }

                    messenger = form.Messenger;

                    settings.RememberContact = form.RememberContact;
                    if ( settings.RememberContact )
                    {
                        settings.RememberContactName( form.Contact );
                    }
                    settings.LastFormat = form.Format == SendFileFormat.Docx ? "DOCX" : "PDF";
                    settings.LastMessenger = messenger == Messenger.Max ? "MAX" : "Telegram";
                    settings.Save( );

                    Cursor.Current = Cursors.WaitCursor;
                    try
                    {
                        string tempFile = saveDocument( form.Format );
                        var options = new SendOptions
                        {
                            Contact = form.Contact ,
                            Message = form.Message ,
                            Format = form.Format ,
                            Messenger = messenger ,
                            OpenByChatName = form.OpenByChatName
                        };

                        if ( messenger == Messenger.Max )
                        {
                            new MaxSender( settings ).SendPreparedFile( options , tempFile );
                        }
                        else
                        {
                            new TelegramSender( settings ).SendPreparedFile( options , tempFile );
                        }
                    }
                    finally
                    {
                        Cursor.Current = Cursors.Default;
                    }
                }

                // Без модального подтверждения: результат виден в самом мессенджере.
                Log( "Отправлено: " + messenger );
            }
            catch ( Exception ex )
            {
                Log( "OnSendButton ERROR: " + ex.ToString( ) );
                MessageBox.Show( "Ошибка: " + ex.Message , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
        }

        public static void SendVia( Settings settings , Func<bool> hasDocument , Func<SendFileFormat,string> saveDocument , string noDocumentMessage , string nativeLabel , bool allowPdf )
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

                if ( allowPdf )
                {
                    using ( var formatForm = new FormatForm( format , nativeLabel ) )
                    {
                        if ( formatForm.ShowDialog( ) != DialogResult.OK )
                        {
                            return;
                        }
                        format = formatForm.Format;
                    }
                }
                else
                {
                    format = SendFileFormat.Docx;
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

                // Без модального подтверждения: окно «Переслать…» открыто в Telegram.
                Log( "Передано в Telegram (окно «Переслать…»)" );
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
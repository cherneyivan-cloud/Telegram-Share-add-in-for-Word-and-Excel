using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace TelegramShareAddin
{
    /// <summary>
    /// Контекст хоста Office, нужный контроллеру для отправки.
    /// Заменяет семь позиционных параметров, которые раньше передавались в SendMessage.
    /// </summary>
    public sealed class SendRequest
    {
        /// <summary>Сообщение, если открытого документа нет.</summary>
        public string NoDocumentMessage { get; set; }

        /// <summary>Название «родного» формата приложения (DOCX/XLSX/MSG) для диалога.</summary>
        public string NativeLabel { get; set; }

        /// <summary>Можно ли предложить PDF.</summary>
        public bool AllowPdf { get; set; }

        /// <summary>Есть ли активный документ.</summary>
        public Func<bool> HasDocument { get; set; }

        /// <summary>Сохранить документ во временный файл (COM, только UI-поток).</summary>
        public Func<SendFileFormat , PreparedFile> SaveDocument { get; set; }
    }

    /// <summary>
    /// Общая (статическая) логика надстроек Word, Excel и Outlook.
    /// COM-интерфейсы и методы ленты реализуются в самих классах надстроек.
    /// </summary>
    internal static class AddinCore
    {
        private static int _activeSend;
        private static IPictureDisp _imageRounded;
        private static IPictureDisp _imagePaperclip;
        private static IPictureDisp _imagePrinter;
        private static IPictureDisp _imageMax;
        private static IPictureDisp _imagePin;
        private static IPictureDisp _imageTelegram;


        public static string GetCustomUi( string ribbonId , bool outlook )
        {
            try
            {
                if ( string.IsNullOrEmpty( ribbonId ) || ribbonId.StartsWith( "Microsoft." , StringComparison.OrdinalIgnoreCase ) )
                {
                    return outlook ? RibbonDefinition.OutlookXml : RibbonDefinition.Xml;
                }
                return null;
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "GetCustomUI" , ex );
                return null;
            }
        }

        public static object GetImage( IRibbonControl control )
        {
            try
            {
                string id = control != null ? control.Id : null;

                if ( id == RibbonIds.SendVia )
                {
                    if ( _imageRounded == null )
                    {
                        _imageRounded = PictureConverter.ToPictureDisp( IconFactory.CreateRounded( ) );
                    }
                    return _imageRounded;
                }

                if ( id == RibbonIds.SendMessage )
                {
                    if ( _imageTelegram == null )
                    {
                        _imageTelegram = PictureConverter.ToPictureDisp( IconFactory.CreateRectangular( ) );
                    }
                    return _imageTelegram;
                }

                if ( id == RibbonIds.Attachments )
                {
                    if ( _imagePaperclip == null )
                    {
                        _imagePaperclip = PictureConverter.ToPictureDisp( IconFactory.CreatePaperclip( ) );
                    }
                    return _imagePaperclip;
                }

                if ( id == RibbonIds.Print )
                {
                    if ( _imagePrinter == null )
                    {
                        _imagePrinter = PictureConverter.ToPictureDisp( IconFactory.CreatePrinter( ) );
                    }
                    return _imagePrinter;
                }

                if ( id == RibbonIds.Max )
                {
                    if ( _imageMax == null )
                    {
                        _imageMax = PictureConverter.ToPictureDisp( IconFactory.CreateMax( ) );
                    }
                    return _imageMax;
                }

                if ( id == RibbonIds.PinQat )
                {
                    if ( _imagePin == null )
                    {
                        _imagePin = PictureConverter.ToPictureDisp( IconFactory.CreatePin( ) );
                    }
                    return _imagePin;
                }
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "GetButtonImage" , ex );
            }
            return null;
        }

        /// <summary>Собирает контекст хоста для диалога и отправки.</summary>
        public static SendRequest CreateRequest( Func<bool> hasDocument , Func<SendFileFormat , PreparedFile> saveDocument ,
                                                 string noDocumentMessage , string nativeLabel , bool allowPdf )
        {
            return new SendRequest
            {
                HasDocument = hasDocument ,
                SaveDocument = saveDocument ,
                NoDocumentMessage = noDocumentMessage ,
                NativeLabel = nativeLabel ,
                AllowPdf = allowPdf
            };
        }

        /// <summary>Кнопки с документом доступны, без документа — нет.</summary>
        public static bool IsEnabled( SendRequest request )
        {
            try
            {
                return request != null && request.HasDocument != null && request.HasDocument( );
            }
            catch
            {
                return false;
            }
        }

        public static void SendMessage( Settings settings , SendRequest request , Messenger messenger )
        {
            if ( !TryEnterSend( ) )
            {
                return;
            }

            try
            {
                AddinLog.Info( "OnSendButton: enter, messenger = " + messenger );

                if ( !IsEnabled( request ) )
                {
                    MessageBox.Show( request.NoDocumentMessage , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                    return;
                }

                Messenger chosen;
                SendFileFormat format;
                string contact;
                string message;
                bool openByChatName;

                using ( var form = new SendForm( settings , request.NativeLabel , request.AllowPdf , messenger ) )
                {
                    if ( form.ShowDialog( ) != DialogResult.OK )
                    {
                        return;
                    }

                    chosen = form.Messenger;
                    format = form.Format;
                    contact = chosen == Messenger.Max ? form.ResolvedContact : form.Contact;
                    message = form.Message;
                    openByChatName = form.OpenByChatName;

                    settings.RememberContact = form.RememberContact;
                    if ( settings.RememberContact )
                    {
                        settings.RememberContactName( form.Contact );
                    }
                    settings.LastFormat = format == SendFileFormat.Native ? "DOCX" : "PDF";
                    if ( !settings.Save( ) )
                    {
                        AddinLog.Warn( "Настройки не сохранены" );
                    }
                }

                // Сохранение документа — COM-вызовы хоста, поэтому только на UI-потоке.
                PreparedFile prepared = request.SaveDocument( format );
                string tempFile = prepared.TempPath;

                var options = new SendOptions
                {
                    Contact = contact ,
                    Message = message ,
                    OpenByChatName = openByChatName ,
                    ClipboardPath = prepared.OriginalPath
                };

                string caption = chosen == Messenger.Max ? "Отправляю в MAX…" : "Отправляю в Telegram…";

                SendRunner.Run( caption , ( ) =>
                {
                    if ( chosen == Messenger.Max )
                    {
                        new MaxSender( settings ).SendPreparedFile( options , tempFile );
                    }
                    else
                    {
                        new TelegramSender( settings ).SendPreparedFile( options , tempFile );
                    }
                } );

                // Без модального подтверждения: результат виден в самом мессенджере.
                AddinLog.Info( "Отправлено: " + chosen );
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "OnSendButton" , ex );
                MessageBox.Show( "Ошибка: " + ex.Message , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
            finally
            {
                Interlocked.Exchange( ref _activeSend , 0 );
            }
        }

        public static void SendVia( Settings settings , SendRequest request )
        {
            if ( !TryEnterSend( ) )
            {
                return;
            }

            try
            {
                AddinLog.Info( "OnSendViaTelegram: enter" );

                if ( !IsEnabled( request ) )
                {
                    MessageBox.Show( request.NoDocumentMessage , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                    return;
                }

                SendFileFormat format = settings.LastFormat == "DOCX" ? SendFileFormat.Native : SendFileFormat.Pdf;

                if ( request.AllowPdf )
                {
                    using ( var formatForm = new FormatForm( format , request.NativeLabel ) )
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
                    format = SendFileFormat.Native;
                }

                PreparedFile prepared = request.SaveDocument( format );

                SendRunner.Run( "Передаю в Telegram…" , ( ) =>
                {
                    new TelegramSender( settings ).SendToActiveChat( prepared.TempPath , prepared.OriginalPath );
                } , true );

                AddinLog.Info( "Передано в Telegram (текущий чат)" );
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "OnSendViaTelegram" , ex );
                MessageBox.Show( "Ошибка: " + ex.Message , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Error );
            }
            finally
            {
                Interlocked.Exchange( ref _activeSend , 0 );
            }
        }

        /// <summary>Защита от повторного запуска: два окна отправки одновременно недопустимы.</summary>
        private static bool TryEnterSend( )
        {
            if ( Interlocked.CompareExchange( ref _activeSend , 1 , 0 ) == 0 )
            {
                return true;
            }

            MessageBox.Show(
                "Отправка уже выполняется — дождитесь её завершения." ,
                "Telegram Share" ,
                MessageBoxButtons.OK ,
                MessageBoxIcon.Information );

            return false;
        }

        public static void Log( string message )
        {
            AddinLog.Info( message );
        }
    }
}

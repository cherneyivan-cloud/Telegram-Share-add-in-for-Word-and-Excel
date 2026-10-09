using System;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace TelegramShareAddin
{
    /// <summary>COM-надстройка для Microsoft Word.</summary>
    [ComVisible( true )]
    [Guid( "AD3B6C2F-4E5F-4A8B-9C1D-3E4F5A6B7C8D" )]
    [ProgId( "TelegramShareAddin" )]
    public class Addin : IDTExtensibility2 , IRibbonExtensibility
    {
        private const string NoDocumentMessage = "Нет открытых документов Word. Откройте документ и повторите.";
        private const string NativeLabel = "DOCX — копия документа";

        private Word.Application _wordApp;
        private Settings _settings;

        // --------------------------------------------------------------
        // IDTExtensibility2
        // --------------------------------------------------------------

        public void OnConnection( object Application , ext_ConnectMode ConnectMode , object AddInInst , ref Array custom )
        {
            try
            {
                AddinCore.Log( "OnConnection: enter" );
                _wordApp = Application as Word.Application;
                _settings = Settings.Load( );
                AddinCore.Log( "OnConnection: exit" );
            }
            catch ( Exception ex )
            {
                AddinCore.Log( "OnConnection ERROR: " + ex.ToString( ) );
            }
        }

        public void OnDisconnection( ext_DisconnectMode RemoveMode , ref Array custom )
        {
            _wordApp = null;
            _settings = null;
        }

        public void OnAddInsUpdate( ref Array custom )
        {
        }

        public void OnStartupComplete( ref Array custom )
        {
        }

        public void OnBeginShutdown( ref Array custom )
        {
        }

        // --------------------------------------------------------------
        // IRibbonExtensibility и обратные вызовы ленты
        // --------------------------------------------------------------

        public string GetCustomUI( string RibbonID )
        {
            return AddinCore.GetCustomUi( RibbonID );
        }

        public object GetButtonImage( IRibbonControl control )
        {
            return AddinCore.GetImage( control );
        }

        public void OnSendButton( IRibbonControl control )
        {
            EnsureSettings( );
            AddinCore.SendMessage( _settings , HasOpenDocument , SaveActiveDocument , NoDocumentMessage , NativeLabel );
        }

        public void OnSendViaTelegram( IRibbonControl control )
        {
            EnsureSettings( );
            AddinCore.SendVia( _settings , HasOpenDocument , SaveActiveDocument , NoDocumentMessage , NativeLabel );
        }

        public void OnOpenGitHub( IRibbonControl control )
        {
            AddinCore.OpenGitHub( );
        }

        // --------------------------------------------------------------

        private void EnsureSettings( )
        {
            if ( _settings == null )
            {
                _settings = Settings.Load( );
            }
        }

        private bool HasOpenDocument( )
        {
            Word.Application app = _wordApp;
            return app != null && app.Documents.Count > 0;
        }

        private string SaveActiveDocument( SendFileFormat format )
        {
            Word.Application app = _wordApp;
            if ( app == null )
            {
                throw new InvalidOperationException( "Word недоступен." );
            }

            Word.Document doc = app.ActiveDocument;

            string originalName = null;
            try
            {
                originalName = doc.Name;
            }
            catch
            {
            }

            string baseName = Path.GetFileNameWithoutExtension( originalName );
            if ( string.IsNullOrEmpty( baseName ) )
            {
                baseName = "Документ";
            }

            string ext;
            if ( format == SendFileFormat.Pdf )
            {
                ext = ".pdf";
            }
            else
            {
                ext = Path.GetExtension( originalName );
                if ( string.IsNullOrEmpty( ext ) )
                {
                    ext = ".docx";
                }
            }

            string path = TelegramSender.NewTempFile( baseName + ext );
            TelegramSender.CleanupOldTempFiles( );

            try
            {
                return WordDocumentSaver.Save( doc , path , format );
            }
            finally
            {
                try
                {
                    Marshal.ReleaseComObject( doc );
                }
                catch
                {
                }
            }
        }
    }
}
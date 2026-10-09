using System;
using System.IO;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace TelegramShareAddin
{
    /// <summary>COM-надстройка для Microsoft Excel.</summary>
    [ComVisible( true )]
    [Guid( "B1C2D3E4-F5A6-4B7C-8D9E-0F1A2B3C4D5E" )]
    [ProgId( "TelegramShareExcel" )]
    public class ExcelAddin : IDTExtensibility2 , IRibbonExtensibility
    {
        private const string NoDocumentMessage = "Нет открытых книг Excel. Откройте книгу и повторите.";
        private const string NativeLabel = "XLSX — копия книги";

        private Excel.Application _excelApp;
        private Settings _settings;

        // --------------------------------------------------------------
        // IDTExtensibility2
        // --------------------------------------------------------------

        public void OnConnection( object Application , ext_ConnectMode ConnectMode , object AddInInst , ref Array custom )
        {
            try
            {
                AddinCore.Log( "OnConnection: enter" );
                _excelApp = Application as Excel.Application;
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
            _excelApp = null;
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
            Excel.Application app = _excelApp;
            return app != null && app.Workbooks.Count > 0;
        }

        private string SaveActiveDocument( SendFileFormat format )
        {
            Excel.Application app = _excelApp;
            if ( app == null )
            {
                throw new InvalidOperationException( "Excel недоступен." );
            }

            Excel.Workbook workbook = app.ActiveWorkbook;

            string originalName = null;
            try
            {
                originalName = workbook.Name;
            }
            catch
            {
            }

            string baseName = Path.GetFileNameWithoutExtension( originalName );
            if ( string.IsNullOrEmpty( baseName ) )
            {
                baseName = "Книга";
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
                    ext = ".xlsx";
                }
            }

            string path = TelegramSender.NewTempFile( baseName + ext );
            TelegramSender.CleanupOldTempFiles( );

            try
            {
                ExcelDocumentSaver.Save( workbook , path , format );
                return path;
            }
            finally
            {
                try
                {
                    Marshal.ReleaseComObject( workbook );
                }
                catch
                {
                }
            }
        }
    }
}
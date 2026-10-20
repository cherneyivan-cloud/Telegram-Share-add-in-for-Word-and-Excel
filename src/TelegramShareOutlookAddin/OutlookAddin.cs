using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TelegramShareAddin
{
    /// <summary>COM-надстройка для Microsoft Outlook.</summary>
    [ComVisible( true )]
    [Guid( "C2D3E4F5-A6B7-4C8D-9E0F-1A2B3C4D5E6F" )]
    [ProgId( "TelegramShareOutlook" )]
    public class OutlookAddin : IDTExtensibility2 , IRibbonExtensibility
    {
        private const string NoDocumentMessage = "Не выбрано письмо Outlook. Выделите письмо или откройте его и повторите.";
        private const string NativeLabel = "MSG — копия письма";

        private Outlook.Application _outlookApp;
        private Settings _settings;

        // --------------------------------------------------------------
        // IDTExtensibility2
        // --------------------------------------------------------------

        public void OnConnection( object Application , ext_ConnectMode ConnectMode , object AddInInst , ref Array custom )
        {
            try
            {
                AddinCore.Log( "OnConnection: enter" );
                _outlookApp = Application as Outlook.Application;
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
            _outlookApp = null;
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
            AddinCore.SendMessage( _settings , HasOpenDocument , SaveActiveDocument , NoDocumentMessage , NativeLabel , false );
        }

        public void OnSendViaTelegram( IRibbonControl control )
        {
            EnsureSettings( );
            AddinCore.SendVia( _settings , HasOpenDocument , SaveActiveDocument , NoDocumentMessage , NativeLabel , false );
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
            Outlook.Application app = _outlookApp;
            if ( app == null )
            {
                return false;
            }

            try
            {
                if ( app.ActiveInspector( ) != null )
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                Outlook.Explorer explorer = app.ActiveExplorer( );
                if ( explorer != null && explorer.Selection != null && explorer.Selection.Count > 0 )
                {
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private string SaveActiveDocument( SendFileFormat format )
        {
            object item = GetActiveItem( );
            if ( item == null )
            {
                throw new InvalidOperationException( NoDocumentMessage );
            }

            try
            {
                return OutlookItemSaver.Save( item , format );
            }
            finally
            {
                try
                {
                    Marshal.ReleaseComObject( item );
                }
                catch
                {
                }
            }
        }

        private object GetActiveItem( )
        {
            Outlook.Application app = _outlookApp;
            if ( app == null )
            {
                return null;
            }

            try
            {
                Outlook.Inspector inspector = app.ActiveInspector( );
                if ( inspector != null )
                {
                    object current = inspector.CurrentItem;
                    if ( current != null )
                    {
                        return current;
                    }
                }
            }
            catch
            {
            }

            try
            {
                Outlook.Explorer explorer = app.ActiveExplorer( );
                if ( explorer != null && explorer.Selection != null && explorer.Selection.Count > 0 )
                {
                    object first = explorer.Selection[ 1 ];
                    if ( first != null )
                    {
                        return first;
                    }
                }
            }
            catch
            {
            }

            return null;
        }
    }
}
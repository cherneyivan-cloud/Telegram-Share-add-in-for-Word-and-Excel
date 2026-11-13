using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Отправка файла в мессенджер MAX.
    ///
    /// Сценарий: открыть чат (по имени из списка или по username через max://),
    /// вставить файл из буфера обмена, при необходимости добавить подпись и нажать
    /// кнопку отправки. Работает через UI Automation, как и Telegram-канал.
    /// </summary>
    internal sealed class MaxSender
    {
        private readonly Settings _settings;

        public MaxSender( Settings settings )
        {
            _settings = settings ?? new Settings( );
        }

        /// <summary>PID окна MAX (0, если MAX не запущен).</summary>
        public static int GetProcessId( )
        {
            IntPtr h = MaxUi.FindWindow( );
            return h == IntPtr.Zero ? 0 : MaxUi.GetProcessId( h );
        }

        public void SendPreparedFile( SendOptions options , string tempFile )
        {
            if ( options == null )
            {
                throw new ArgumentNullException( "options" );
            }

            string contact = ( options.Contact ?? string.Empty ).Trim( );
            if ( contact.Length == 0 )
            {
                throw new ArgumentException( "Не указан контакт." , "Contact" );
            }

            try
            {
                IntPtr hwnd = MaxUi.EnsureWindow( 15000 );
                MaxUi.Activate( hwnd );
                Thread.Sleep( 700 );

                int pid = MaxUi.GetProcessId( hwnd );
                bool opened = false;

                if ( options.OpenByChatName )
                {
                    opened = MaxUi.OpenChat( pid , contact , 8000 );
                }

                // Ссылку max://max.ru/<username> используем только для явно указанного
                // username (начинается с «@»). Иначе MAX открывает страницу
                // «Ссылка устарела или не работает» на произвольный текст.
                if ( !opened && contact.StartsWith( "@" , StringComparison.Ordinal ) )
                {
                    string username = MaxUi.NormalizeUsername( contact );
                    if ( username != null && MaxUi.OpenByUsername( username ) )
                    {
                        Thread.Sleep( _settings.ChatOpenDelayMs );
                        hwnd = MaxUi.EnsureWindow( 15000 );
                        MaxUi.Activate( hwnd );
                        pid = MaxUi.GetProcessId( hwnd );
                        opened = true;
                    }
                }

                if ( !opened )
                {
                    opened = MaxUi.OpenChat( pid , contact , 8000 );
                }

                if ( !opened )
                {
                    throw new InvalidOperationException(
                        "Не удалось найти чат «" + contact + "» в MAX. Выберите чат из списка (кнопка «Обновить из MAX») " +
                        "или откройте его в MAX вручную и повторите отправку." );
                }

                Thread.Sleep( Math.Max( 800 , _settings.ChatOpenDelayMs / 2 ) );

                MaxUi.FocusComposer( hwnd );
                Thread.Sleep( 300 );

                ClipboardHelper.SetFileAndPath( tempFile );
                SendKeys.SendWait( "^v" );

                if ( !MaxUi.WaitForPendingAttachment( hwnd , 12000 ) )
                {
                    throw new InvalidOperationException( "MAX не принял файл из буфера обмена. Откройте чат в MAX и повторите." );
                }

                if ( !string.IsNullOrEmpty( options.Message ) )
                {
                    ClipboardHelper.SetText( options.Message );
                    SendKeys.SendWait( "^v" );
                    Thread.Sleep( 600 );
                }

                if ( !MaxUi.Send( hwnd , 12000 ) )
                {
                    throw new InvalidOperationException( "Не удалось отправить файл в MAX. Проверьте окно MAX и повторите." );
                }

                Thread.Sleep( 1000 );
            }
            finally
            {
                // После сохранения полный путь к файлу остаётся в буфере обмена.
                try
                {
                    ClipboardHelper.SetText( tempFile );
                }
                catch
                {
                }
            }
        }
    }

    /// <summary>UI Automation для окна MAX.</summary>
    internal static class MaxUi
    {
        private const int SW_RESTORE = 9;
        private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const int MOUSEEVENTF_LEFTUP = 0x0004;
        private const int MOUSEEVENTF_WHEEL = 0x0800;

        public static IEnumerable<Process> GetProcesses( )
        {
            return Process.GetProcesses( ).Where(
                p => string.Equals( p.ProcessName , "max" , StringComparison.OrdinalIgnoreCase ) );
        }

        public static IntPtr FindWindow( )
        {
            foreach ( Process p in GetProcesses( ) )
            {
                try
                {
                    IntPtr h = p.MainWindowHandle;
                    if ( h != IntPtr.Zero && IsWindowVisible( h ) )
                    {
                        return h;
                    }
                }
                catch
                {
                }
            }
            return IntPtr.Zero;
        }

        public static IntPtr EnsureWindow( int timeoutMs )
        {
            int deadline = Environment.TickCount + timeoutMs;
            while ( Environment.TickCount < deadline )
            {
                IntPtr h = FindWindow( );
                if ( h != IntPtr.Zero )
                {
                    return h;
                }
                Thread.Sleep( 300 );
            }

            throw new InvalidOperationException( "Не найдено окно MAX. Запустите MAX и повторите." );
        }

        public static void Activate( IntPtr hwnd )
        {
            try
            {
                ShowWindow( hwnd , SW_RESTORE );
                SetForegroundWindow( hwnd );
            }
            catch
            {
            }
        }

        public static int GetProcessId( IntPtr hwnd )
        {
            int pid;
            GetWindowThreadProcessId( hwnd , out pid );
            return pid;
        }

        /// <summary>
        /// Username MAX из строки вида «@name». Возвращает null, если это не похоже
        /// на username (кириллица, пробелы, отображаемое имя чата и т.п.).
        /// </summary>
        public static string NormalizeUsername( string contact )
        {
            string c = ( contact ?? string.Empty ).Trim( );
            if ( !c.StartsWith( "@" , StringComparison.Ordinal ) )
            {
                return null;
            }

            c = c.Substring( 1 ).Trim( ).ToLowerInvariant( );

            if ( c.Length < 2 || c.Length > 40 )
            {
                return null;
            }

            foreach ( char ch in c )
            {
                bool ok = ( ch >= 'a' && ch <= 'z' ) ||
                          ( ch >= '0' && ch <= '9' ) ||
                          ch == '_' || ch == '.' || ch == '-';
                if ( !ok )
                {
                    return null;
                }
            }

            if ( c.All( char.IsDigit ) )
            {
                return null;
            }

            return c;
        }

        public static bool OpenByUsername( string username )
        {
            try
            {
                string uri = "max://max.ru/" + Uri.EscapeDataString( username );
                Process.Start( new ProcessStartInfo( uri ) { UseShellExecute = true } );
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static List<MessengerChatInfo> GetChats( int pid )
        {
            var result = new List<MessengerChatInfo>( );
            var seen = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

            foreach ( AutomationElement item in ChatButtons( pid ) )
            {
                string name = SafeName( item );
                if ( !string.IsNullOrEmpty( name ) && seen.Add( name ) )
                {
                    result.Add( new MessengerChatInfo { Name = name , Type = string.Empty } );
                }
            }

            return result;
        }

        public static bool OpenChat( int pid , string chatName , int timeoutMs )
        {
            if ( pid == 0 || string.IsNullOrEmpty( chatName ) )
            {
                return false;
            }

            int deadline = Environment.TickCount + timeoutMs;

            while ( Environment.TickCount < deadline )
            {
                AutomationElement item = FindChat( pid , chatName );
                if ( item != null && ClickElement( item ) )
                {
                    return true;
                }

                ScrollChatList( pid , -4 );
                Thread.Sleep( 400 );

                item = FindChat( pid , chatName );
                if ( item != null && ClickElement( item ) )
                {
                    return true;
                }

                if ( Environment.TickCount >= deadline )
                {
                    break;
                }

                ScrollChatList( pid , 4 );
                Thread.Sleep( 400 );
            }

            return false;
        }

        public static void FocusComposer( IntPtr hwnd )
        {
            AutomationElement button = RightmostSendButton( hwnd );
            if ( button == null )
            {
                return;
            }

            System.Windows.Rect buttonRect;
            System.Windows.Rect windowRect;
            if ( !TryGetRect( button , out buttonRect ) || !GetWindowRectSafe( hwnd , out windowRect ) )
            {
                return;
            }

            int x = (int)( buttonRect.Left - 300 );
            int y = (int)( buttonRect.Top + buttonRect.Height / 2 );

            if ( x < (int)windowRect.Left + 20 )
            {
                x = (int)windowRect.Left + 20;
            }

            ClickAt( x , y );
        }

        public static bool WaitForPendingAttachment( IntPtr hwnd , int timeoutMs )
        {
            int pid = GetProcessId( hwnd );
            int deadline = Environment.TickCount + timeoutMs;

            while ( Environment.TickCount < deadline )
            {
                if ( HasPendingAttachment( pid ) )
                {
                    return true;
                }
                Thread.Sleep( 300 );
            }

            return false;
        }

        /// <summary>Нажать кнопку отправки. true — вложение принято (превью исчезло).</summary>
        public static bool Send( IntPtr hwnd , int timeoutMs )
        {
            int pid = GetProcessId( hwnd );
            int deadline = Environment.TickCount + timeoutMs;
            int attempts = 0;

            while ( Environment.TickCount < deadline )
            {
                if ( !HasPendingAttachment( pid ) )
                {
                    return true;
                }

                if ( attempts < 5 )
                {
                    AutomationElement button = RightmostSendButton( hwnd );
                    if ( button != null && ClickElement( button ) )
                    {
                        attempts++;
                        Thread.Sleep( 900 );
                        continue;
                    }
                }

                Thread.Sleep( 300 );
            }

            return !HasPendingAttachment( pid );
        }

        // --------------------------------------------------------------

        private static AutomationElement FindChat( int pid , string chatName )
        {
            foreach ( AutomationElement item in ChatButtons( pid ) )
            {
                if ( string.Equals( SafeName( item ) , chatName.Trim( ) , StringComparison.OrdinalIgnoreCase ) )
                {
                    return item;
                }
            }
            return null;
        }

        private static List<AutomationElement> ChatButtons( int pid )
        {
            var result = new List<AutomationElement>( );

            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                try
                {
                    var condition = new AndCondition(
                        new PropertyCondition( AutomationElement.AutomationIdProperty , "ChatDelegate" ) ,
                        new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.Button ) );

                    AutomationElementCollection found = window.FindAll( TreeScope.Descendants , condition );
                    foreach ( AutomationElement e in found )
                    {
                        result.Add( e );
                    }
                }
                catch
                {
                }
            }

            return result;
        }

        private static void ScrollChatList( int pid , int steps )
        {
            List<AutomationElement> buttons = ChatButtons( pid );
            if ( buttons.Count == 0 )
            {
                return;
            }

            System.Windows.Rect r;
            if ( !TryGetRect( buttons[0] , out r ) )
            {
                return;
            }

            SetCursorPos( (int)( r.Left + 30 ) , (int)( r.Top + 40 ) );
            Thread.Sleep( 60 );

            int delta = steps >= 0 ? -120 : 120;
            for ( int i = 0; i < Math.Abs( steps ); i++ )
            {
                mouse_event( MOUSEEVENTF_WHEEL , 0 , 0 , delta , UIntPtr.Zero );
                Thread.Sleep( 80 );
            }

            Thread.Sleep( 250 );
        }

        private static AutomationElement RightmostSendButton( IntPtr hwnd )
        {
            int pid = GetProcessId( hwnd );
            AutomationElement best = null;
            double bestLeft = double.MinValue;

            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                try
                {
                    AutomationElementCollection found = window.FindAll(
                        TreeScope.Descendants ,
                        new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.Button ) );

                    foreach ( AutomationElement e in found )
                    {
                        string id = SafeAutomationId( e );
                        if ( id.IndexOf( "SendPanelButton" , StringComparison.OrdinalIgnoreCase ) < 0 )
                        {
                            continue;
                        }

                        System.Windows.Rect r;
                        if ( !TryGetRect( e , out r ) )
                        {
                            continue;
                        }

                        if ( r.Left > bestLeft )
                        {
                            bestLeft = r.Left;
                            best = e;
                        }
                    }
                }
                catch
                {
                }
            }

            return best;
        }

        private static bool HasPendingAttachment( int pid )
        {
            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                try
                {
                    AutomationElementCollection all = window.FindAll(
                        TreeScope.Descendants , Condition.TrueCondition );

                    foreach ( AutomationElement e in all )
                    {
                        string id = SafeAutomationId( e );
                        if ( id.IndexOf( "FileSendAttachDelegate" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                             id.IndexOf( "AttachRemoveButton" , StringComparison.OrdinalIgnoreCase ) >= 0 )
                        {
                            return true;
                        }
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static AutomationElementCollection TopWindows( int pid )
        {
            var condition = new PropertyCondition( AutomationElement.ProcessIdProperty , pid );
            return AutomationElement.RootElement.FindAll( TreeScope.Children , condition );
        }

        private static bool ClickElement( AutomationElement element )
        {
            try
            {
                object pattern;
                if ( element.TryGetCurrentPattern( InvokePattern.Pattern , out pattern ) )
                {
                    ( (InvokePattern)pattern ).Invoke( );
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                object pattern;
                if ( element.TryGetCurrentPattern( SelectionItemPattern.Pattern , out pattern ) )
                {
                    ( (SelectionItemPattern)pattern ).Select( );
                    return true;
                }
            }
            catch
            {
            }

            System.Windows.Rect r;
            if ( !TryGetRect( element , out r ) )
            {
                return false;
            }

            return ClickAt( (int)( r.Left + r.Width / 2 ) , (int)( r.Top + r.Height / 2 ) );
        }

        private static bool ClickAt( int x , int y )
        {
            try
            {
                SetCursorPos( x , y );
                Thread.Sleep( 50 );
                mouse_event( MOUSEEVENTF_LEFTDOWN , 0 , 0 , 0 , UIntPtr.Zero );
                mouse_event( MOUSEEVENTF_LEFTUP , 0 , 0 , 0 , UIntPtr.Zero );
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetRect( AutomationElement element , out System.Windows.Rect rect )
        {
            rect = System.Windows.Rect.Empty;
            try
            {
                rect = element.Current.BoundingRectangle;
                return !rect.IsEmpty && rect.Width > 0 && rect.Height > 0;
            }
            catch
            {
                return false;
            }
        }

        private static string SafeName( AutomationElement element )
        {
            try
            {
                return element.Current.Name;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeAutomationId( AutomationElement element )
        {
            try
            {
                return element.Current.AutomationId ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool GetWindowRectSafe( IntPtr hwnd , out System.Windows.Rect rect )
        {
            rect = System.Windows.Rect.Empty;
            try
            {
                RECT r;
                if ( !GetWindowRect( hwnd , out r ) )
                {
                    return false;
                }

                rect = new System.Windows.Rect( r.Left , r.Top , r.Right - r.Left , r.Bottom - r.Top );
                return true;
            }
            catch
            {
                return false;
            }
        }

        [System.Runtime.InteropServices.StructLayout( System.Runtime.InteropServices.LayoutKind.Sequential )]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool SetForegroundWindow( IntPtr hWnd );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool ShowWindow( IntPtr hWnd , int nCmdShow );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool IsWindowVisible( IntPtr hWnd );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern int GetWindowThreadProcessId( IntPtr hWnd , out int processId );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool SetCursorPos( int x , int y );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern void mouse_event( int dwFlags , int dx , int dy , int dwData , UIntPtr dwExtraInfo );

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool GetWindowRect( IntPtr hWnd , out RECT rect );
    }
}

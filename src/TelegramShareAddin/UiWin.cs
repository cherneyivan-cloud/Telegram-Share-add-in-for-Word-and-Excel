using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Общие операции с окнами Windows и защищённая отправка клавиш.
    ///
    /// Клавиши уходят в активное окно, поэтому перед каждой вставкой проверяется,
    /// что на переднем плане именно нужный процесс. Иначе отправка отклоняется —
    /// лучше ошибка, чем текст, вставленный в чужое приложение.
    /// </summary>
    internal static class UiWin
    {
        public const int SwRestore = 9;

        public const uint MouseLeftDown = 0x0002;
        public const uint MouseLeftUp = 0x0004;
        public const uint MouseRightDown = 0x0008;
        public const uint MouseRightUp = 0x0010;
        public const uint MouseWheel = 0x0800;

        public static IntPtr GetForeground( )
        {
            return GetForegroundWindow( );
        }

        public static int ProcessIdOf( IntPtr hwnd )
        {
            if ( hwnd == IntPtr.Zero )
            {
                return 0;
            }

            uint pid;
            GetWindowThreadProcessId( hwnd , out pid );
            return ( int ) pid;
        }

        /// <summary>true, если на переднем плане окно процесса.</summary>
        public static bool IsForeground( int processId )
        {
            return processId != 0 && ProcessIdOf( GetForegroundWindow( ) ) == processId;
        }

        /// <summary>Разворачивает и выводит окно на передний план; true — окно действительно активно.</summary>
        public static bool EnsureForeground( IntPtr hwnd , int attempts = 8 , int delayMs = 150 )
        {
            if ( hwnd == IntPtr.Zero )
            {
                return false;
            }

            try
            {
                if ( IsIconic( hwnd ) )
                {
                    ShowWindow( hwnd , SwRestore );
                }
            }
            catch
            {
            }

            for ( int i = 0 ; i < attempts ; i++ )
            {
                if ( GetForegroundWindow( ) == hwnd )
                {
                    return true;
                }

                SetForegroundWindow( hwnd );
                BringWindowToTop( hwnd );
                Thread.Sleep( delayMs );
            }

            return GetForegroundWindow( ) == hwnd;
        }

        /// <summary>Отправить клавиши в окно процесса. Бросает исключение, если фокус ушёл.</summary>
        public static void SendKeysTo( int processId , string keys , string action )
        {
            if ( !IsForeground( processId ) )
            {
                throw new InvalidOperationException(
                    "Окно " + action + " не на переднем плане — отправка отменена. " +
                    "Не переключайте окна во время отправки и повторите." );
            }

            SendKeys.SendWait( keys );
        }

        /// <summary>Все окна процесса верхнего уровня (для UI Automation).</summary>
        public static AutomationElementCollection TopWindows( int processId )
        {
            var condition = new PropertyCondition( AutomationElement.ProcessIdProperty , processId );
            return AutomationElement.RootElement.FindAll( TreeScope.Children , condition );
        }

        public static void ClickAt( int x , int y )
        {
            SetCursorPos( x , y );
            Thread.Sleep( 60 );
            mouse_event( MouseLeftDown , 0 , 0 , 0 , UIntPtr.Zero );
            mouse_event( MouseLeftUp , 0 , 0 , 0 , UIntPtr.Zero );
        }

        /// <summary>Щелчок по центру элемента; false — элемент без геометрии.</summary>
        public static bool ClickElement( AutomationElement element )
        {
            try
            {
                System.Windows.Rect rect = element.Current.BoundingRectangle;
                if ( rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0 )
                {
                    return false;
                }

                ClickAt( ( int ) ( rect.Left + rect.Width / 2 ) , ( int ) ( rect.Top + rect.Height / 2 ) );
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>Ждёт выполнения условия с шагом опроса; возвращает false по таймауту.</summary>
        public static bool WaitUntil( Func<bool> condition , int timeoutMs , int stepMs = 150 )
        {
            int waited = 0;

            while ( waited < timeoutMs )
            {
                bool ready;
                try
                {
                    ready = condition( );
                }
                catch
                {
                    ready = false;
                }

                if ( ready )
                {
                    return true;
                }

                Thread.Sleep( stepMs );
                waited += stepMs;
            }

            return false;
        }

        /// <summary>Процессы по имени без утечки handle: возвращает снимок Id/пути.</summary>
        public static Process[ ] FindProcesses( string processNamePart )
        {
            Process[ ] all;
            try
            {
                all = Process.GetProcesses( );
            }
            catch
            {
                return new Process[ 0 ];
            }

            var result = new System.Collections.Generic.List<Process>( );

            foreach ( Process process in all )
            {
                bool match;
                try
                {
                    match = process.ProcessName.IndexOf( processNamePart , StringComparison.OrdinalIgnoreCase ) >= 0;
                }
                catch
                {
                    match = false;
                }

                if ( match )
                {
                    result.Add( process );
                }
                else
                {
                    process.Dispose( );
                }
            }

            return result.ToArray( );
        }

        /// <summary>Снимок процесса: без RCW/Handle, Disposing не требуется.</summary>
        public sealed class WindowProcess
        {
            public int Id;
            public IntPtr MainWindowHandle;
            public string ExecutablePath;
        }

        /// <summary>Снимок процессов по имени: объекты Process сразу освобождаются.</summary>
        public static System.Collections.Generic.List<WindowProcess> SnapshotProcesses( string processNamePart )
        {
            var result = new System.Collections.Generic.List<WindowProcess>( );

            foreach ( Process process in FindProcesses( processNamePart ) )
            {
                try
                {
                    var info = new WindowProcess { Id = process.Id };

                    try { info.MainWindowHandle = process.MainWindowHandle; }
                    catch { }

                    try { info.ExecutablePath = process.MainModule.FileName; }
                    catch { }

                    result.Add( info );
                }
                catch
                {
                }
                finally
                {
                    process.Dispose( );
                }
            }

            return result;
        }

        [StructLayout( LayoutKind.Sequential )]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport( "user32.dll" )]
        private static extern IntPtr GetForegroundWindow( );

        [DllImport( "user32.dll" )]
        public static extern bool SetForegroundWindow( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        public static extern bool BringWindowToTop( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        public static extern bool ShowWindow( IntPtr hWnd , int nCmdShow );

        [DllImport( "user32.dll" )]
        public static extern bool IsIconic( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        public static extern bool IsWindowVisible( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        public static extern uint GetWindowThreadProcessId( IntPtr hWnd , out uint processId );

        [DllImport( "user32.dll" )]
        public static extern bool GetCursorPos( out POINT point );

        [DllImport( "user32.dll" )]
        public static extern bool SetCursorPos( int x , int y );

        [DllImport( "user32.dll" )]
        public static extern void mouse_event( uint flags , uint dx , uint dy , uint data , UIntPtr extraInfo );

        [DllImport( "user32.dll" )]
        public static extern void keybd_event( byte vk , byte scan , uint flags , UIntPtr extraInfo );

        [DllImport( "user32.dll" )]
        public static extern IntPtr WindowFromPoint( POINT point );

        [DllImport( "user32.dll" )]
        public static extern IntPtr GetAncestor( IntPtr hWnd , uint flags );
    }
}

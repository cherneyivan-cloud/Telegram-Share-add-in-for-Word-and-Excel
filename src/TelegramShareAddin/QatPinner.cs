using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Добавление кнопок надстройки на панель быстрого доступа в заголовке окна Office.
    /// В разметке ленты надстройки элемент qat недоступен: Office разрешает его только
    /// при startFromScratch="true", то есть с полной заменой ленты. Поэтому используется
    /// то же действие, что и пункт контекстного меню «Добавить на панель быстрого доступа».
    /// </summary>
    internal static class QatPinner
    {
        private const string TabLabel = "Отправить";

        private static readonly string[ ] ButtonLabels = { "Отправить в Telegram" , "Отправить в MAX" };

        private static int _running;

        /// <summary>Запускает добавление кнопок в фоновом потоке, чтобы не блокировать окно Office.</summary>
        public static void Start( )
        {
            if ( Interlocked.CompareExchange( ref _running , 1 , 0 ) != 0 )
            {
                return;
            }

            IntPtr foreground = GetForegroundWindow( );
            ThreadPool.QueueUserWorkItem( state => Run( (IntPtr) state ) , foreground );
        }

        private static void Run( IntPtr foregroundWindow )
        {
            try
            {
                AddinCore.Log( "QatPinner: enter" );

                AutomationElement window = FindWindow( foregroundWindow );
                if ( window == null )
                {
                    AddinCore.Log( "QatPinner: окно приложения не найдено" );
                    Fail( );
                    return;
                }

                AutomationElement tab = WaitFor( window , ControlType.TabItem , TabLabel , 2000 );
                if ( tab == null )
                {
                    AddinCore.Log( "QatPinner: вкладка не найдена" );
                    Fail( );
                    return;
                }

                IntPtr windowHandle = new IntPtr( window.Current.NativeWindowHandle );
                EnsureForeground( windowHandle );

                EnsureSelected( tab );
                Thread.Sleep( 900 );

                int added = 0;
                int already = 0;

                foreach ( string label in ButtonLabels )
                {
                    AutomationElement button = WaitFor( window , ControlType.Button , label , 1500 );
                    if ( button == null )
                    {
                        // Лента может быть свёрнута — разворачиваем (Ctrl+F1) и пробуем снова.
                        SendCtrlF1( );
                        Thread.Sleep( 700 );
                        button = WaitFor( window , ControlType.Button , label , 2000 );
                    }

                    if ( button == null )
                    {
                        AddinCore.Log( "QatPinner: кнопка не найдена: " + label );
                        continue;
                    }

                    if ( CountVisible( window , label ) > 1 )
                    {
                        already++;
                        AddinCore.Log( "QatPinner: уже закреплена: " + label );
                        continue;
                    }

                    if ( PinOne( window , label ) )
                    {
                        added++;
                    }
                }

                int pinned = 0;
                foreach ( string label in ButtonLabels )
                {
                    if ( CountVisible( window , label ) > 1 )
                    {
                        pinned++;
                    }
                }

                AddinCore.Log( "QatPinner: добавлено=" + added + ", уже было=" + already + ", в заголовке=" + pinned );
                Report( pinned , already );
            }
            catch ( Exception ex )
            {
                AddinCore.Log( "QatPinner ERROR: " + ex.ToString( ) );
                Fail( );
            }
            finally
            {
                Interlocked.Exchange( ref _running , 0 );
            }
        }

        private static bool PinOne( AutomationElement window , string label )
        {
            string stage = "поиск кнопки";
            POINT saved = new POINT( );
            bool cursorSaved = false;

            for ( int attempt = 0 ; attempt < 2 ; attempt++ )
            {
                try
                {
                    stage = "поиск кнопки";
                    // Элемент перечитываем заново: после переключения вкладки прежний может стать недоступным.
                    AutomationElement button = WaitFor( window , ControlType.Button , label , 2000 );
                    if ( button == null )
                    {
                        return false;
                    }

                    stage = "координаты кнопки";
                    System.Windows.Rect rect = button.Current.BoundingRectangle;
                    if ( rect.Width < 2 || rect.Height < 2 )
                    {
                        return false;
                    }

                    int x = ( int ) ( rect.X + rect.Width / 2 );
                    int y = ( int ) ( rect.Y + rect.Height / 2 );

                    stage = "наведение";
                    if ( !EnsureForeground( new IntPtr( window.Current.NativeWindowHandle ) ) )
                    {
                        AddinCore.Log( "QatPinner: окно не на переднем плане: " + label );
                        return false;
                    }
                    Thread.Sleep( 250 );

                    GetCursorPos( out saved );
                    cursorSaved = true;

                    SetCursorPos( x , y );
                    Thread.Sleep( 200 );

                    POINT point = new POINT( );
                    point.X = x;
                    point.Y = y;
                    IntPtr under = WindowFromPoint( point );
                    IntPtr rootWindow = under == IntPtr.Zero ? IntPtr.Zero : GetAncestor( under , GaRoot );
                    if ( rootWindow != new IntPtr( window.Current.NativeWindowHandle ) )
                    {
                        AddinCore.Log( "QatPinner: под курсором не окно Office: " + label );
                        SetCursorPos( saved.X , saved.Y );
                        cursorSaved = false;
                        return false;
                    }

                    mouse_event( RightDown , 0 , 0 , 0 , UIntPtr.Zero );
                    mouse_event( RightUp , 0 , 0 , 0 , UIntPtr.Zero );
                    Thread.Sleep( 900 );

                    stage = "поиск пункта меню";
                    AutomationElement item = WaitForAddToQat( window , 2500 );

                    if ( item == null )
                    {
                        AddinCore.Log( "QatPinner: пункт меню не найден для " + label );
                        PressEscape( );
                        return false;
                    }

                    // Курсор возвращаем только после вызова: если увести его раньше,
                    // всплывающее меню Office успевает закрыться.
                    stage = "вызов пункта меню";
                    AddinCore.Log( "QatPinner: пункт меню '" + SafeName( item ) + "'" );

                    if ( !InvokeItem( item ) )
                    {
                        AddinCore.Log( "QatPinner: пункт меню не сработал для " + label );
                        PressEscape( );
                        SetCursorPos( saved.X , saved.Y );
                        cursorSaved = false;
                        return false;
                    }

                    SetCursorPos( saved.X , saved.Y );
                    cursorSaved = false;
                    Thread.Sleep( 600 );
                    return true;
                }
                catch ( Exception ex )
                {
                    AddinCore.Log( "QatPinner: попытка " + ( attempt + 1 ) + " не удалась для " + label +
                                   " (" + ex.GetType( ).Name + " на этапе «" + stage + "»: " + ex.Message + ")" );
                    if ( cursorSaved )
                    {
                        SetCursorPos( saved.X , saved.Y );
                        cursorSaved = false;
                    }
                    if ( attempt == 0 )
                    {
                        Thread.Sleep( 800 );
                        continue;
                    }
                    return false;
                }
            }

            return false;
        }

        private static string SafeName( AutomationElement element )
        {
            try
            {
                return element.Current.Name;
            }
            catch
            {
                return "?";
            }
        }

        /// <summary>Вызывает пункт меню; если шаблон недоступен — щёлкает по нему.</summary>
        private static bool InvokeItem( AutomationElement item )
        {
            try
            {
                var pattern = item.GetCurrentPattern( InvokePattern.Pattern ) as InvokePattern;
                if ( pattern != null )
                {
                    pattern.Invoke( );
                    return true;
                }
            }
            catch ( Exception ex )
            {
                AddinCore.Log( "QatPinner: Invoke не сработал (" + ex.GetType( ).Name + ": " + ex.Message + ")" );
            }

            try
            {
                System.Windows.Rect rect = item.Current.BoundingRectangle;
                if ( rect.Width < 2 || rect.Height < 2 )
                {
                    return false;
                }

                SetCursorPos( ( int ) ( rect.X + rect.Width / 2 ) , ( int ) ( rect.Y + rect.Height / 2 ) );
                Thread.Sleep( 150 );
                mouse_event( LeftDown , 0 , 0 , 0 , UIntPtr.Zero );
                mouse_event( LeftUp , 0 , 0 , 0 , UIntPtr.Zero );
                return true;
            }
            catch ( Exception ex )
            {
                AddinCore.Log( "QatPinner: щелчок по пункту меню не удался (" + ex.Message + ")" );
                return false;
            }
        }

        private static AutomationElement WaitForAddToQat( AutomationElement window , int timeoutMs )
        {
            int waited = 0;

            while ( waited < timeoutMs )
            {
                var condition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.MenuItem );
                AutomationElementCollection items = window.FindAll( TreeScope.Descendants , condition );

                foreach ( AutomationElement item in items )
                {
                    string name;

                    try
                    {
                        name = item.Current.Name;
                    }
                    catch
                    {
                        continue;
                    }

                    if ( string.IsNullOrEmpty( name ) )
                    {
                        continue;
                    }

                    // Только пункт именно про панель быстрого доступа: остальные пункты меню
                    // (папки, поиск, уведомления) трогать нельзя.
                    bool addVerb = name.StartsWith( "Добавить" , StringComparison.OrdinalIgnoreCase ) ||
                                   name.StartsWith( "Add" , StringComparison.OrdinalIgnoreCase ) ||
                                   name.StartsWith( "Añadir" , StringComparison.OrdinalIgnoreCase ) ||
                                   name.StartsWith( "Hinzufügen" , StringComparison.OrdinalIgnoreCase );

                    bool quickAccess = name.IndexOf( "быстрого доступа" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                                       name.IndexOf( "Quick Access" , StringComparison.OrdinalIgnoreCase ) >= 0;

                    if ( addVerb && quickAccess )
                    {
                        return item;
                    }
                }

                Thread.Sleep( 150 );
                waited += 150;
            }

            return null;
        }

        private static int CountVisible( AutomationElement window , string label )
        {
            var condition = new AndCondition(
                new PropertyCondition( AutomationElement.NameProperty , label ) ,
                new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.Button ) );

            int count = 0;
            foreach ( AutomationElement item in window.FindAll( TreeScope.Descendants , condition ) )
            {
                try
                {
                    if ( item.Current.BoundingRectangle.Width > 2 )
                    {
                        count++;
                    }
                }
                catch
                {
                }
            }
            return count;
        }

        private static AutomationElement WaitFor( AutomationElement parent , ControlType type , string name , int timeoutMs )
        {
            int waited = 0;
            while ( waited < timeoutMs )
            {
                AutomationElement found = FindFirst( parent , type , name );
                if ( found != null )
                {
                    return found;
                }
                Thread.Sleep( 150 );
                waited += 150;
            }
            return null;
        }

        private static AutomationElement FindFirst( AutomationElement parent , ControlType type , string name )
        {
            var condition = new AndCondition(
                new PropertyCondition( AutomationElement.NameProperty , name ) ,
                new PropertyCondition( AutomationElement.ControlTypeProperty , type ) );
            return parent.FindFirst( TreeScope.Descendants , condition );
        }

        private static void EnsureSelected( AutomationElement tab )
        {
            try
            {
                var selection = tab.GetCurrentPattern( SelectionItemPattern.Pattern ) as SelectionItemPattern;
                if ( selection != null && !selection.Current.IsSelected )
                {
                    selection.Select( );
                }
            }
            catch ( Exception ex )
            {
                AddinCore.Log( "QatPinner: не удалось выбрать вкладку: " + ex.Message );
            }
        }

        private static AutomationElement FindWindow( IntPtr foregroundWindow )
        {
            uint pid = ( uint ) Process.GetCurrentProcess( ).Id;

            if ( foregroundWindow != IntPtr.Zero )
            {
                uint windowPid;
                GetWindowThreadProcessId( foregroundWindow , out windowPid );
                if ( windowPid == pid )
                {
                    try
                    {
                        return AutomationElement.FromHandle( foregroundWindow );
                    }
                    catch
                    {
                    }
                }
            }

            AutomationElement root = AutomationElement.RootElement;
            var condition = new PropertyCondition( AutomationElement.ProcessIdProperty , ( int ) pid );
            AutomationElement best = null;
            double bestArea = 0;

            foreach ( AutomationElement candidate in root.FindAll( TreeScope.Children , condition ) )
            {
                try
                {
                    System.Windows.Rect rect = candidate.Current.BoundingRectangle;
                    double area = rect.Width * rect.Height;
                    if ( area > bestArea )
                    {
                        bestArea = area;
                        best = candidate;
                    }
                }
                catch
                {
                }
            }
            return best;
        }

        private static void Report( int pinned , int already )
        {
            string text;

            if ( pinned >= ButtonLabels.Length )
            {
                text = already >= ButtonLabels.Length
                    ? "Кнопки «Отправить в Telegram» и «Отправить в MAX» уже находятся в заголовке окна."
                    : "Кнопки «Отправить в Telegram» и «Отправить в MAX» добавлены в заголовок окна." +
                      Environment.NewLine + Environment.NewLine +
                      "Дальше они всегда будут под рукой — рядом с кнопкой «Сохранить».";
            }
            else if ( pinned > 0 )
            {
                text = "Одна кнопка уже в заголовке окна, вторую добавить не удалось." + Environment.NewLine + Environment.NewLine +
                       "Щёлкните правой кнопкой мыши по нужной кнопке на вкладке «Отправить» " +
                       "и выберите «Добавить на панель быстрого доступа».";
            }
            else
            {
                text = "Не удалось добавить кнопки автоматически." + Environment.NewLine + Environment.NewLine +
                       "Щёлкните правой кнопкой мыши по кнопке «Отправить в Telegram» на вкладке «Отправить» " +
                       "и выберите «Добавить на панель быстрого доступа».";
            }

            MessageBox.Show( text , "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Information );
        }

        private static void Fail( )
        {
            MessageBox.Show(
                "Не удалось добавить кнопки в заголовок окна." + Environment.NewLine + Environment.NewLine +
                "Щёлкните правой кнопкой мыши по кнопке «Отправить в Telegram» на вкладке «Отправить» " +
                "и выберите «Добавить на панель быстрого доступа»." ,
                "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Warning );
        }

        /// <summary>Разворачивает окно при необходимости и выводит его на передний план.</summary>
        private static bool EnsureForeground( IntPtr hWnd )
        {
            if ( hWnd == IntPtr.Zero )
            {
                return false;
            }

            try
            {
                if ( IsIconic( hWnd ) )
                {
                    ShowWindow( hWnd , SwRestore );
                }
            }
            catch
            {
            }

            for ( int i = 0 ; i < 8 ; i++ )
            {
                if ( GetForegroundWindow( ) == hWnd )
                {
                    return true;
                }

                SetForegroundWindow( hWnd );
                BringWindowToTop( hWnd );
                Thread.Sleep( 150 );
            }

            return GetForegroundWindow( ) == hWnd;
        }

        private static void PressEscape( )
        {
            keybd_event( VkEscape , 0 , 0 , UIntPtr.Zero );
            keybd_event( VkEscape , 0 , KeyUp , UIntPtr.Zero );
        }

        private static void SendCtrlF1( )
        {
            keybd_event( VkControl , 0 , 0 , UIntPtr.Zero );
            keybd_event( VkF1 , 0 , 0 , UIntPtr.Zero );
            keybd_event( VkF1 , 0 , KeyUp , UIntPtr.Zero );
            keybd_event( VkControl , 0 , KeyUp , UIntPtr.Zero );
        }

        // --------------------------------------------------------------
        // Win32
        // --------------------------------------------------------------

        private const uint RightDown = 0x0008;
        private const uint RightUp = 0x0010;
        private const uint LeftDown = 0x0002;
        private const uint LeftUp = 0x0004;
        private const uint KeyUp = 0x0002;
        private const byte VkControl = 0x11;
        private const byte VkEscape = 0x1B;
        private const byte VkF1 = 0x70;
        private const int SwRestore = 9;
        private const uint GaRoot = 2;

        [StructLayout( LayoutKind.Sequential )]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport( "user32.dll" )]
        private static extern IntPtr GetForegroundWindow( );

        [DllImport( "user32.dll" )]
        private static extern bool SetForegroundWindow( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        private static extern uint GetWindowThreadProcessId( IntPtr hWnd , out uint processId );

        [DllImport( "user32.dll" )]
        private static extern bool GetCursorPos( out POINT point );

        [DllImport( "user32.dll" )]
        private static extern bool SetCursorPos( int x , int y );

        [DllImport( "user32.dll" )]
        private static extern void mouse_event( uint flags , uint dx , uint dy , uint data , UIntPtr extraInfo );

        [DllImport( "user32.dll" )]
        private static extern void keybd_event( byte vk , byte scan , uint flags , UIntPtr extraInfo );

        [DllImport( "user32.dll" )]
        private static extern bool ShowWindow( IntPtr hWnd , int nCmdShow );

        [DllImport( "user32.dll" )]
        private static extern bool IsIconic( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        private static extern bool BringWindowToTop( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        private static extern IntPtr WindowFromPoint( POINT point );

        [DllImport( "user32.dll" )]
        private static extern IntPtr GetAncestor( IntPtr hWnd , uint flags );
    }
}

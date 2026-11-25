using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
                IntPtr hwnd = MaxUi.EnsureWindow( Timeouts.WindowWaitMs );
                UiWin.EnsureForeground( hwnd );
                Thread.Sleep( 700 );

                int pid = MaxUi.GetProcessId( hwnd );
                bool opened = false;

                if ( options.OpenByChatName )
                {
                    opened = MaxUi.OpenChat( pid , contact , Timeouts.ChatOpenMs );
                }

                // Ссылку max://max.ru/<username> используем только для явно указанного
                // username (начинается с «@»). Иначе MAX открывает страницу
                // «Ссылка устарела или не работает» на произвольный текст.
                if ( !opened && contact.StartsWith( "@" , StringComparison.Ordinal ) )
                {
                    string username = MaxUi.NormalizeUsername( contact );
                    if ( username != null && MaxUi.OpenByUsername( username ) )
                    {
                        Thread.Sleep( _settings.ChatOpenDelay );
                        hwnd = MaxUi.EnsureWindow( Timeouts.WindowWaitMs );
                        UiWin.EnsureForeground( hwnd );
                        pid = MaxUi.GetProcessId( hwnd );
                        opened = true;
                    }
                }

                if ( !opened )
                {
                    opened = MaxUi.OpenChat( pid , contact , Timeouts.ChatOpenMs );
                }

                if ( !opened )
                {
                    throw new InvalidOperationException(
                        "Не удалось найти чат «" + contact + "» в MAX. Выберите чат из списка (кнопка «Обновить из MAX») " +
                        "или откройте его в MAX вручную и повторите отправку." );
                }

                Thread.Sleep( Math.Max( 800 , _settings.ChatOpenDelay / 2 ) );

                MaxUi.FocusComposer( hwnd );
                Thread.Sleep( 300 );

                ClipboardHelper.SetFileAndPath( tempFile );

                // Перед вставкой убеждаемся, что активен именно MAX.
                if ( !UiWin.EnsureForeground( hwnd ) )
                {
                    throw new InvalidOperationException( "Не удалось вывести MAX на передний план — отправка отменена." );
                }

                UiWin.SendKeysTo( pid , "^v" , "MAX" );

                if ( !MaxUi.WaitForPendingAttachment( hwnd , Timeouts.AttachmentMs ) )
                {
                    throw new InvalidOperationException( "MAX не принял файл из буфера обмена. Откройте чат в MAX и повторите." );
                }

                if ( !string.IsNullOrEmpty( options.Message ) )
                {
                    ClipboardHelper.SetText( options.Message );
                    UiWin.SendKeysTo( pid , "^v" , "MAX" );
                    Thread.Sleep( 600 );
                }

                if ( !MaxUi.Send( hwnd , Timeouts.AttachmentMs ) )
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

        public static List<UiWin.WindowProcess> GetProcesses( )
        {
            return UiWin.SnapshotProcesses( "max" );
        }

        private static string FindMaxExe( )
        {
            foreach ( UiWin.WindowProcess p in GetProcesses( ) )
            {
                if ( !string.IsNullOrEmpty( p.ExecutablePath ) && File.Exists( p.ExecutablePath ) )
                {
                    return p.ExecutablePath;
                }
            }

            string[ ] candidates =
            {
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) , @"MAX\max.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData ) , @"MAX\max.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ProgramFiles ) , @"MAX\max.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ProgramFilesX86 ) , @"MAX\max.exe" )
            };

            foreach ( string c in candidates )
            {
                if ( File.Exists( c ) )
                {
                    return c;
                }
            }

            return null;
        }

        public static IntPtr FindWindow( )
        {
            IntPtr best = IntPtr.Zero;
            int bestArea = 0;

            foreach ( UiWin.WindowProcess p in GetProcesses( ) )
            {
                int pid = p.Id;

                EnumWindows( delegate( IntPtr h , IntPtr l )
                {
                    int wpid;
                    GetWindowThreadProcessId( h , out wpid );
                    if ( wpid != pid )
                    {
                        return true;
                    }

                    RECT r;
                    if ( !GetWindowRect( h , out r ) )
                    {
                        return true;
                    }

                    string title = GetWindowTitle( h );
                    int area = ( r.Right - r.Left ) * ( r.Bottom - r.Top );

                    // Главное окно MAX — с заголовком «MAX» и нормальным размером.
                    // Служебные окна («max», «QTrayIconMessageWindow», «_q_titlebar»)
                    // исключаем по размеру/заголовку.
                    if ( string.Equals( title , "MAX" , StringComparison.OrdinalIgnoreCase ) && area > 200000 )
                    {
                        if ( area > bestArea )
                        {
                            bestArea = area;
                            best = h;
                        }
                    }

                    return true;
                } , IntPtr.Zero );
            }

            if ( best != IntPtr.Zero )
            {
                return best;
            }

            // Запасной вариант: обычное видимое главное окно.
            foreach ( UiWin.WindowProcess p in GetProcesses( ) )
            {
                if ( p.MainWindowHandle != IntPtr.Zero )
                {
                    return p.MainWindowHandle;
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

        private static List<MessengerChatInfo> _cachedChats;
        private static DateTime _cacheTime;
        private static int _cachePid;

        /// <summary>Кэш списка чатов: повторное открытие диалога не дёргает MAX.</summary>
        public static List<MessengerChatInfo> GetChats( int pid )
        {
            return GetChats( pid , false , CancellationToken.None );
        }

        public static List<MessengerChatInfo> GetChats( int pid , bool refresh )
        {
            return GetChats( pid , refresh , CancellationToken.None );
        }

        public static List<MessengerChatInfo> GetChats( int pid , bool refresh , CancellationToken token )
        {
            if ( !refresh && _cachedChats != null && _cachePid == pid &&
                 ( DateTime.UtcNow - _cacheTime ).TotalMinutes < 5 )
            {
                return new List<MessengerChatInfo>( _cachedChats );
            }

            List<MessengerChatInfo> chats = ReadChats( pid , token );
            if ( chats.Count > 0 )
            {
                _cachedChats = chats;
                _cachePid = pid;
                _cacheTime = DateTime.UtcNow;
            }

            return chats;
        }

        private static List<MessengerChatInfo> ReadChats( int pid , CancellationToken token )
        {
            var result = new List<MessengerChatInfo>( );
            var seen = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

            IntPtr hwnd = FindWindow( );
            if ( hwnd == IntPtr.Zero )
            {
                return result;
            }

            // Основной (детерминированный) способ: у скроллбара списка чатов MAX
            // есть RangeValuePattern (0..100). Двигаем его по позициям и собираем
            // видимые чаты. Это быстро и не зависит от фокуса/переднего плана.
            AutomationElement scrollBar = FindChatScrollBar( pid );
            object rangeObject;
            if ( scrollBar != null &&
                 scrollBar.TryGetCurrentPattern( RangeValuePattern.Pattern , out rangeObject ) )
            {
                var range = (RangeValuePattern)rangeObject;

                try { range.SetValue( 0 ); } catch { }

                for ( double v = 0 ; v <= 100 ; v += 5 )
                {
                    if ( token.IsCancellationRequested )
                    {
                        return result;
                    }

                    try { range.SetValue( v ); } catch { }
                    Thread.Sleep( 60 );

                    foreach ( string name in ReadVisibleNames( pid ) )
                    {
                        if ( seen.Add( name ) )
                        {
                            result.Add( new MessengerChatInfo { Name = name , Type = string.Empty } );
                        }
                    }
                }

                try { range.SetValue( 0 ); } catch { }
                return result;
            }

            // Запасной вариант (нет RangeValuePattern) — прокрутка колесом.
            // Здесь окно должно быть активным: Qt направляет колесо активному окну.
            UiWin.EnsureForeground( hwnd );
            Thread.Sleep( 300 );

            ScrollToTop( hwnd );

            int added = CollectChats( hwnd , result , seen , token );

            // При «холодном» старте первая прокрутка может не сработать. Повторяем.
            if ( added < 15 )
            {
                Thread.Sleep( 600 );
                result.Clear( );
                seen.Clear( );
                CollectChats( hwnd , result , seen , token );
            }

            // Возвращаем список к началу.
            ScrollToTop( hwnd );
            return result;
        }

        /// <summary>Скроллбар списка чатов MAX (поддерживает RangeValuePattern).</summary>
        private static AutomationElement FindChatScrollBar( int pid )
        {
            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                try
                {
                    var condition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.ScrollBar );
                    AutomationElementCollection bars = window.FindAll( TreeScope.Descendants , condition );

                    foreach ( AutomationElement bar in bars )
                    {
                        string id = SafeAutomationId( bar );
                        if ( id.IndexOf( "LeadingChatsBar" , StringComparison.OrdinalIgnoreCase ) >= 0 &&
                             id.IndexOf( "MeScrollBar" , StringComparison.OrdinalIgnoreCase ) >= 0 )
                        {
                            return bar;
                        }
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        /// <summary>Листает список вниз и добавляет в result/seen новые чаты.</summary>
        private static int CollectChats( IntPtr hwnd , List<MessengerChatInfo> result , HashSet<string> seen , CancellationToken token )
        {
            int pid = GetProcessId( hwnd );
            int before = result.Count;

            ScrollToTop( hwnd );

            string previousSignature = null;
            int unchangedPasses = 0;
            const int MaxPasses = 400;

            for ( int pass = 0; pass < MaxPasses; pass++ )
            {
                if ( token.IsCancellationRequested )
                {
                    break;
                }

                List<string> visibleNames = ReadVisibleNames( pid );
                string signature = string.Join( "\u001f" , visibleNames );

                foreach ( string name in visibleNames )
                {
                    if ( seen.Add( name ) )
                    {
                        result.Add( new MessengerChatInfo { Name = name , Type = string.Empty } );
                    }
                }

                if ( string.Equals( signature , previousSignature , StringComparison.Ordinal ) )
                {
                    unchangedPasses++;
                }
                else
                {
                    unchangedPasses = 0;
                    previousSignature = signature;
                }

                if ( visibleNames.Count == 0 || unchangedPasses >= 2 )
                {
                    break;
                }

                ScrollChatList( hwnd , 10 );
                Thread.Sleep( 130 );
            }

            return result.Count - before;
        }

        public static bool OpenChat( int pid , string chatName , int timeoutMs )
        {
            if ( pid == 0 || string.IsNullOrEmpty( chatName ) )
            {
                return false;
            }

            IntPtr hwnd = FindWindow( );

            int deadline = Environment.TickCount + timeoutMs;

            while ( Environment.TickCount < deadline )
            {
                AutomationElement item = FindChat( pid , chatName );
                if ( item == null )
                {
                    item = FindChatBySubstring( pid , chatName );
                }
                if ( item != null && ClickElement( item ) )
                {
                    return true;
                }

                ScrollChatList( hwnd , -4 );
                Thread.Sleep( 400 );

                item = FindChat( pid , chatName );
                if ( item == null )
                {
                    item = FindChatBySubstring( pid , chatName );
                }
                if ( item != null && ClickElement( item ) )
                {
                    return true;
                }

                if ( Environment.TickCount >= deadline )
                {
                    break;
                }

                ScrollChatList( hwnd , 4 );
                Thread.Sleep( 400 );
            }

            return false;
        }

        public static void FocusComposer( IntPtr hwnd )
        {
            // Основной путь — поле ввода, найденное через UI Automation.
            AutomationElement composer = FindComposer( hwnd );
            if ( composer != null )
            {
                try
                {
                    composer.SetFocus( );
                    AddinLog.Info( "MAX: фокус в поле ввода через UIA" );
                    return;
                }
                catch ( Exception ex )
                {
                    AddinLog.Warn( "MAX: SetFocus на поле ввода не сработал — " + AddinLog.Describe( ex ) );
                }

                if ( UiWin.ClickElement( composer ) )
                {
                    return;
                }
            }

            // Резервный путь: клик слева от кнопки отправки.
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

            AddinLog.Warn( "MAX: поле ввода не найдено, кликаю по резервной точке" );

            int x = (int)( buttonRect.Left - 300 );
            int y = (int)( buttonRect.Top + buttonRect.Height / 2 );

            if ( x < (int)windowRect.Left + 20 )
            {
                x = (int)windowRect.Left + 20;
            }

            ClickAt( x , y );
        }

        /// <summary>Поле ввода сообщения: Edit/Document в нижней половине окна.</summary>
        private static AutomationElement FindComposer( IntPtr hwnd )
        {
            int pid = GetProcessId( hwnd );
            if ( pid == 0 )
            {
                return null;
            }

            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                System.Windows.Rect windowRect;
                if ( !TryGetRect( window , out windowRect ) )
                {
                    continue;
                }

                double middle = windowRect.Top + windowRect.Height / 2;

                foreach ( ControlType type in new[ ] { ControlType.Edit , ControlType.Document } )
                {
                    var condition = new PropertyCondition( AutomationElement.ControlTypeProperty , type );
                    AutomationElementCollection candidates = window.FindAll( TreeScope.Descendants , condition );

                    AutomationElement best = null;
                    double widest = 0;

                    foreach ( AutomationElement candidate in candidates )
                    {
                        System.Windows.Rect rect;
                        if ( !TryGetRect( candidate , out rect ) )
                        {
                            continue;
                        }

                        if ( rect.Top < middle || rect.Width < 200 || rect.Height < 20 )
                        {
                            continue;
                        }

                        if ( rect.Width > widest )
                        {
                            widest = rect.Width;
                            best = candidate;
                        }
                    }

                    if ( best != null )
                    {
                        return best;
                    }
                }
            }

            return null;
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

        /// <summary>Ищет чат, имя которого содержит заданную подстроку (только если однозначно).</summary>
        private static AutomationElement FindChatBySubstring( int pid , string chatName )
        {
            string needle = ( chatName ?? string.Empty ).Trim( );
            if ( string.IsNullOrEmpty( needle ) )
            {
                return null;
            }

            AutomationElement match = null;
            foreach ( AutomationElement item in ChatButtons( pid ) )
            {
                string name = SafeName( item );
                if ( string.IsNullOrEmpty( name ) )
                {
                    continue;
                }

                if ( name.IndexOf( needle , StringComparison.OrdinalIgnoreCase ) >= 0 )
                {
                    if ( match != null )
                    {
                        return null; // неоднозначно
                    }
                    match = item;
                }
            }

            return match;
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

        private static void ScrollChatList( IntPtr hwnd , int steps )
        {
            if ( steps == 0 )
            {
                return;
            }

            int pid = GetProcessId( hwnd );

            AutomationElement chatBar = FindChatBar( pid );
            if ( chatBar == null )
            {
                return;
            }

            System.Windows.Rect r;
            if ( !TryGetRect( chatBar , out r ) )
            {
                return;
            }

            // Прокручиваем колесом (mouse_event) по одному «щелчку» за раз —
            // крупные дельты Qt обрабатывает непредсказуемо. Курсор предварительно
            // ставим над списком: Qt направляет прокрутку элементу под курсором.
            int x = (int)( r.Left + r.Width / 2 );
            int y = (int)( r.Top + r.Height - 180 );
            SetCursorPos( x , y );
            Thread.Sleep( 30 );

            int delta = steps > 0 ? -120 : 120;
            for ( int i = 0; i < Math.Abs( steps ); i++ )
            {
                mouse_event( MOUSEEVENTF_WHEEL , 0 , 0 , delta , UIntPtr.Zero );
                Thread.Sleep( 15 );
            }

            Thread.Sleep( 160 );
        }

        /// <summary>Имена видимых (экспонированных UIA) чатов в порядке списка.</summary>
        private static List<string> ReadVisibleNames( int pid )
        {
            var names = new List<string>( );
            foreach ( AutomationElement item in ChatButtons( pid ) )
            {
                string name = SafeName( item );
                if ( !string.IsNullOrEmpty( name ) )
                {
                    names.Add( name );
                }
            }
            return names;
        }

        /// <summary>Прокрутка к началу списка (до стабилизации первого элемента).</summary>
        private static void ScrollToTop( IntPtr hwnd )
        {
            int pid = GetProcessId( hwnd );
            string previousFirst = null;
            int unchanged = 0;
            int emptyPasses = 0;

            for ( int i = 0; i < 60; i++ )
            {
                List<string> names = ReadVisibleNames( pid );
                string first = names.Count > 0 ? names[0] : string.Empty;

                if ( names.Count == 0 )
                {
                    emptyPasses++;
                    if ( emptyPasses >= 3 )
                    {
                        break;
                    }
                }
                else
                {
                    emptyPasses = 0;
                }

                if ( string.Equals( first , previousFirst , StringComparison.Ordinal ) )
                {
                    unchanged++;
                }
                else
                {
                    unchanged = 0;
                    previousFirst = first;
                }

                if ( names.Count > 0 && unchanged >= 3 )
                {
                    break;
                }

                ScrollChatList( hwnd , -20 );
                Thread.Sleep( 120 );
            }
        }

        private static AutomationElement FindChatBar( int pid )
        {
            foreach ( AutomationElement window in TopWindows( pid ) )
            {
                try
                {
                    AutomationElement bar = window.FindFirst(
                        TreeScope.Descendants,
                        new PropertyCondition( AutomationElement.AutomationIdProperty , "LeadingChatsBar" ) );
                    if ( bar != null )
                    {
                        return bar;
                    }
                }
                catch
                {
                }
            }

            return null;
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

        [System.Runtime.InteropServices.DllImport( "user32.dll" )]
        private static extern bool EnumWindows( EnumWindowsProc callback , IntPtr lParam );

        [System.Runtime.InteropServices.DllImport( "user32.dll" , CharSet = System.Runtime.InteropServices.CharSet.Unicode )]
        private static extern int GetWindowText( IntPtr hWnd , System.Text.StringBuilder text , int maxCount );

        private delegate bool EnumWindowsProc( IntPtr hWnd , IntPtr lParam );

        private static string GetWindowTitle( IntPtr hWnd )
        {
            try
            {
                var sb = new System.Text.StringBuilder( 256 );
                GetWindowText( hWnd , sb , sb.Capacity );
                return sb.ToString( );
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}

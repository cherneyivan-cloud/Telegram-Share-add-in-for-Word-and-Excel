using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>Формат временного файла.</summary>
    public enum SendFileFormat
    {
        Pdf = 0,
        Docx = 1
    }

    /// <summary>Параметры отправки.</summary>
    public class TelegramSendOptions
    {
        public string Contact { get; set; }
        public string Message { get; set; }
        public SendFileFormat Format { get; set; }

        /// <summary>true — контакт выбран из списка чатов Telegram (открываем чат кликом(.</summary>
        public bool OpenByChatName { get; set; }
    }

    /// <summary>Информация о чате Telegram.</summary>
    public class TelegramChatInfo
    {
        public string Name { get; set; }
        public string Type { get; set; }

        public bool IsGroup
        {
            get
            {
                if ( string.IsNullOrEmpty( Type ) )
                {
                    return false;
                }
                return Type.IndexOf( "Группа" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                       Type.IndexOf( "Group" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                       Type.IndexOf( "Канал" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                       Type.IndexOf( "Channel" , StringComparison.OrdinalIgnoreCase ) >= 0;
            }
        }
    }

    /// <summary>
    /// Работа с Telegram Desktop: буфер обмена, открытие чата, вставка файла,
    /// подтверждение отправки. Не зависит от конкретного приложения Office.
    /// </summary>
    public class TelegramSender
    {
        private const int SW_RESTORE = 9;
        private const int TempFileLifetimeMs = 10 * 60 * 1000;

        private readonly Settings _settings;

        public TelegramSender( Settings settings )
        {
            _settings = settings ?? new Settings( );
        }

        /// <summary>Путь во временной папке с заданным именем файла.</summary>
        public static string NewTempFile( string fileName )
        {
            string dir = Path.Combine( Path.GetTempPath( ) , "TelegramShare" );
            Directory.CreateDirectory( dir );
            return Path.Combine( dir , SanitizeFileName( fileName ) );
        }

        private static string SanitizeFileName( string name )
        {
            if ( string.IsNullOrEmpty( name ) )
            {
                return "document";
            }

            foreach ( char c in Path.GetInvalidFileNameChars( ) )
            {
                name = name.Replace( c , '_' );
            }
            return name;
        }

        /// <summary>Удалить временные файлы прошлых отправок (старше суток).</summary>
        public static void CleanupOldTempFiles( )
        {
            DeleteOld( Path.Combine( Path.GetTempPath( ) , "TelegramShare" ) );
            DeleteOld( Path.GetTempPath( ) );
        }

        private static void DeleteOld( string dir )
        {
            try
            {
                if ( !Directory.Exists( dir ) )
                {
                    return;
                }
                foreach ( string file in Directory.GetFiles( dir , "TelegramShare*" ) )
                {
                    try
                    {
                        if ( File.GetLastWriteTime( file ) < DateTime.Now.AddDays( -1 ) )
                        {
                            File.Delete( file );
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        /// <summary>Отправить уже подготовленный файл выбранному контакту (автоматически(.</summary>
        public void SendPreparedFile( TelegramSendOptions options , string tempFile )
        {
            if ( string.IsNullOrWhiteSpace( options.Contact ) )
            {
                throw new ArgumentException( "Не указан контакт." , "Contact" );
            }

            IDataObject savedClipboard = null;
            try
            {
                try
                {
                    savedClipboard = Clipboard.GetDataObject( );
                }
                catch
                {
                    savedClipboard = null;
                }

                SetFileToClipboard( tempFile );
                OpenChatAndSend( options );
                ScheduleDelete( tempFile , TempFileLifetimeMs );
            }
            finally
            {
                RestoreClipboard( savedClipboard );
            }
        }

        /// <summary>Открыть в Telegram окно «Переслать…» для уже подготовленного файла.</summary>
        public void SendWithPickerPreparedFile( string tempFile )
        {
            LaunchTelegramSendPath( tempFile );
            ScheduleDelete( tempFile , TempFileLifetimeMs );
        }

        private static void SetFileToClipboard( string filePath )
        {
            var files = new StringCollection( );
            files.Add( filePath );

            try
            {
                Clipboard.SetFileDropList( files );
            }
            catch ( ExternalException ex )
            {
                throw new InvalidOperationException( "Не удалось поместить файл в буфер обмена: " + ex.Message , ex );
            }
        }

        private void OpenChatAndSend( TelegramSendOptions options )
        {
            int telegramPid;

            if ( options.OpenByChatName )
            {
                telegramPid = GetTelegramProcessId( );
                if ( telegramPid == 0 )
                {
                    throw new InvalidOperationException( "Telegram Desktop не запущен. Запустите Telegram и повторите." );
                }

                IntPtr hwnd = FindTelegramWindow( );
                if ( hwnd == IntPtr.Zero )
                {
                    throw new InvalidOperationException( "Не найдено окно Telegram Desktop. Разверните окно Telegram из трея." );
                }

                ShowWindow( hwnd , SW_RESTORE );
                SetForegroundWindow( hwnd );
                Thread.Sleep( 400 );

                bool opened = TelegramUi.OpenChat( telegramPid , options.Contact.Trim( ) , 8000 );
                if ( !opened )
                {
                    throw new InvalidOperationException( "Не удалось открыть чат «" + options.Contact + "» в Telegram. Нажмите «Обновить из Telegram» и выберите чат заново." );
                }

                Thread.Sleep( _settings.ChatOpenDelayMs );
            }
            else
            {
                string uri = BuildTelegramUri( options.Contact.Trim( ) );
                LaunchTelegram( uri );
                Thread.Sleep( _settings.ChatOpenDelayMs );

                IntPtr hwnd = FindTelegramWindow( );
                if ( hwnd == IntPtr.Zero )
                {
                    throw new InvalidOperationException( "Не найдено окно Telegram Desktop. Разверните окно Telegram из трея и повторите." );
                }

                ShowWindow( hwnd , SW_RESTORE );
                SetForegroundWindow( hwnd );
                Thread.Sleep( 400 );

                telegramPid = GetTelegramProcessId( );
            }

            SendKeys.SendWait( "^v" );
            AddinCore.Log( "Send: файл вставлен, ожидаю диалог Telegram" );
            Thread.Sleep( _settings.PasteDelayMs );

            bool confirmed = TelegramUi.ConfirmSend( options.Message , telegramPid , 8000 );
            AddinCore.Log( "Send: UIA подтверждение = " + confirmed );

            if ( !confirmed )
            {
                if ( !string.IsNullOrEmpty( options.Message ) )
                {
                    SendKeys.SendWait( EscapeSendKeys( options.Message ) );
                    Thread.Sleep( 300 );
                }
                SendKeys.SendWait( "{ENTER}" );
            }

            Thread.Sleep( 1500 );
        }

        private static void LaunchTelegram( string uri )
        {
            try
            {
                Process.Start( new ProcessStartInfo( uri ) { UseShellExecute = true , Verb = "open" } );
            }
            catch ( Exception ex )
            {
                throw new InvalidOperationException( "Не удалось открыть Telegram. Убедитесь, что Telegram Desktop установлен." , ex );
            }
        }

        private static void LaunchTelegramSendPath( string file )
        {
            string exe = FindTelegramExe( );
            if ( string.IsNullOrEmpty( exe ) )
            {
                throw new InvalidOperationException( "Не найден Telegram Desktop (Telegram.exe). Установите Telegram или запустите его один раз." );
            }

            try
            {
                Process.Start( new ProcessStartInfo( exe , "-sendpath \"" + file + "\"" ) { UseShellExecute = false } );
            }
            catch ( Exception ex )
            {
                throw new InvalidOperationException( "Не удалось открыть Telegram с файлом: " + ex.Message , ex );
            }
        }

        private static string FindTelegramExe( )
        {
            foreach ( var p in GetTelegramProcesses( ) )
            {
                try
                {
                    string path = p.MainModule.FileName;
                    if ( !string.IsNullOrEmpty( path ) && File.Exists( path ) )
                    {
                        return path;
                    }
                }
                catch
                {
                }
            }

            string[ ] candidates = new string[ ]
            {
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ApplicationData ) , @"Telegram Desktop\Telegram.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData ) , @"Programs\Telegram Desktop\Telegram.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ProgramFiles ) , @"Telegram Desktop\Telegram.exe" ) ,
                Path.Combine( Environment.GetFolderPath( Environment.SpecialFolder.ProgramFilesX86 ) , @"Telegram Desktop\Telegram.exe" )
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

        private static string BuildTelegramUri( string contact )
        {
            string c = contact.Trim( );

            if ( c.StartsWith( "@" , StringComparison.Ordinal ) )
            {
                c = c.Substring( 1 );
            }

            if ( string.IsNullOrEmpty( c ) )
            {
                throw new ArgumentException( "Некорректный контакт." , "Contact" );
            }

            if ( c.StartsWith( "+" , StringComparison.Ordinal ) )
            {
                string digits = new string( c.Where( char.IsDigit ).ToArray( ) );
                if ( digits.Length < 5 )
                {
                    throw new ArgumentException( "Некорректный номер телефона." , "Contact" );
                }
                return "tg://resolve?phone=" + "+" + digits;
            }

            if ( c.All( char.IsDigit ) )
            {
                if ( c.Length >= 10 )
                {
                    return "tg://resolve?phone=" + c;
                }
                return "tg://user?id=" + c;
            }

            return "tg://resolve?domain=" + Uri.EscapeDataString( c );
        }

        private static IntPtr FindTelegramWindow( )
        {
            foreach ( var p in GetTelegramProcesses( ) )
            {
                IntPtr h = p.MainWindowHandle;
                if ( h != IntPtr.Zero && IsWindowVisible( h ) )
                {
                    return h;
                }
            }
            return IntPtr.Zero;
        }

        public static int GetTelegramProcessId( )
        {
            foreach ( var p in GetTelegramProcesses( ) )
            {
                if ( p.MainWindowHandle != IntPtr.Zero )
                {
                    return p.Id;
                }
            }
            return 0;
        }

        private static IEnumerable<Process> GetTelegramProcesses( )
        {
            return Process.GetProcesses( ).Where( x => x.ProcessName.IndexOf( "Telegram" , StringComparison.OrdinalIgnoreCase ) >= 0 );
        }

        private static string EscapeSendKeys( string text )
        {
            if ( string.IsNullOrEmpty( text ) )
            {
                return text;
            }

            var sb = new StringBuilder( text.Length + 16 );
            foreach ( char ch in text )
            {
                switch ( ch )
                {
                    case '+' : sb.Append( "{+}" ); break;
                    case '^' : sb.Append( "{^}" ); break;
                    case '%' : sb.Append( "{%}" ); break;
                    case '~' : sb.Append( "{~}" ); break;
                    case '(' : sb.Append( "{(}" ); break;
                    case ')' : sb.Append( "{)}" ); break;
                    case '[' : sb.Append( "{[}" ); break;
                    case ']' : sb.Append( "{]}" ); break;
                    case '{' : sb.Append( "{{}" ); break;
                    case '}' : sb.Append( "{}}" ); break;
                    default : sb.Append( ch ); break;
                }
            }
            return sb.ToString( );
        }

        private static void RestoreClipboard( IDataObject saved )
        {
            try
            {
                if ( saved != null )
                {
                    Clipboard.SetDataObject( saved , true );
                }
                else
                {
                    Clipboard.Clear( );
                }
            }
            catch
            {
            }
        }

        private static void ScheduleDelete( string path , int delayMs )
        {
            ThreadPool.QueueUserWorkItem( delegate
            {
                try
                {
                    Thread.Sleep( delayMs );
                    if ( File.Exists( path ) )
                    {
                        File.Delete( path );
                    }
                }
                catch
                {
                }
            } );
        }

        [DllImport( "user32.dll" )]
        private static extern bool SetForegroundWindow( IntPtr hWnd );

        [DllImport( "user32.dll" )]
        private static extern bool ShowWindow( IntPtr hWnd , int nCmdShow );

        [DllImport( "user32.dll" )]
        private static extern bool IsWindowVisible( IntPtr hWnd );
    }

    /// <summary>
    /// Работа с интерфейсом Telegram Desktop через UI Automation.
    /// </summary>
    internal static class TelegramUi
    {
        private sealed class ChatEntry
        {
            public AutomationElement Element;
            public string Name;
            public string Type;
            public string OwnName;
        }

        public static List<TelegramChatInfo> GetChats( int processId )
        {
            var result = new List<TelegramChatInfo>( );
            if ( processId == 0 )
            {
                return result;
            }

            AutomationElement chatList = FindFirstChatList( processId );
            if ( chatList == null )
            {
                return result;
            }

            List<ChatEntry> entries = ReadChatEntries( chatList );
            var seen = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

            foreach ( ChatEntry e in entries )
            {
                if ( string.IsNullOrEmpty( e.Name ) )
                {
                    continue;
                }
                if ( IsFolder( e.Type ) || IsFolder( e.OwnName ) )
                {
                    continue;
                }
                if ( seen.Add( e.Name ) )
                {
                    result.Add( new TelegramChatInfo { Name = e.Name , Type = e.Type } );
                }
            }

            return result;
        }

        public static bool OpenChat( int processId , string chatName , int timeoutMs )
        {
            if ( processId == 0 || string.IsNullOrEmpty( chatName ) )
            {
                return false;
            }

            int deadline = Environment.TickCount + timeoutMs;
            while ( Environment.TickCount < deadline )
            {
                try
                {
                    AutomationElement chatList = FindFirstChatList( processId );
                    if ( chatList != null )
                    {
                        List<ChatEntry> entries = ReadChatEntries( chatList );
                        foreach ( ChatEntry e in entries )
                        {
                            if ( string.Equals( e.Name , chatName , StringComparison.OrdinalIgnoreCase ) )
                            {
                                if ( SelectChatItem( e.Element ) )
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                catch
                {
                }

                Thread.Sleep( 300 );
            }

            return false;
        }

        public static bool ConfirmSend( string caption , int processId , int timeoutMs )
        {
            if ( processId == 0 )
            {
                return false;
            }

            int deadline = Environment.TickCount + timeoutMs;
            while ( Environment.TickCount < deadline )
            {
                try
                {
                    foreach ( AutomationElement window in TopWindows( processId ) )
                    {
                        AutomationElement button = FindSendButton( window );
                        if ( button != null )
                        {
                            if ( !string.IsNullOrEmpty( caption ) )
                            {
                                SetCaption( window , caption );
                            }
                            if ( InvokeElement( button ) )
                            {
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                }

                Thread.Sleep( 300 );
            }

            return false;
        }

        private static List<ChatEntry> ReadChatEntries( AutomationElement chatList )
        {
            var entries = new List<ChatEntry>( );

            var cache = new CacheRequest( );
            cache.Add( AutomationElement.NameProperty );
            cache.Add( AutomationElement.ControlTypeProperty );
            cache.Add( ValuePattern.ValueProperty );
            cache.TreeScope = TreeScope.Subtree;

            using ( cache.Activate( ) )
            {
                var itemCondition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.ListItem );
                AutomationElementCollection items = chatList.FindAll( TreeScope.Children , itemCondition );

                foreach ( AutomationElement item in items )
                {
                    var entry = new ChatEntry( );
                    entry.Element = item;
                    entry.OwnName = GetCachedString( item , AutomationElement.NameProperty );

                    var dataCondition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.DataItem );
                    AutomationElementCollection data = item.FindAll( TreeScope.Children , dataCondition );

                    foreach ( AutomationElement d in data )
                    {
                        string field = GetCachedString( d , AutomationElement.NameProperty );
                        if ( string.IsNullOrEmpty( field ) )
                        {
                            continue;
                        }

                        if ( field == "Тип" || field == "Type" )
                        {
                            entry.Type = GetCachedValue( d );
                        }
                        else if ( field == "Название" || field == "Title" || field == "Name" )
                        {
                            entry.Name = GetCachedValue( d );
                        }
                    }

                    entries.Add( entry );
                }
            }

            if ( entries.Count == 0 )
            {
                entries = ReadChatEntriesLive( chatList );
            }

            return entries;
        }

        private static List<ChatEntry> ReadChatEntriesLive( AutomationElement chatList )
        {
            var entries = new List<ChatEntry>( );
            var itemCondition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.ListItem );
            AutomationElementCollection items = chatList.FindAll( TreeScope.Children , itemCondition );

            var nameCondition = new AndCondition(
                new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.DataItem ) ,
                new OrCondition(
                    new PropertyCondition( AutomationElement.NameProperty , "Название" ) ,
                    new PropertyCondition( AutomationElement.NameProperty , "Title" ) ,
                    new PropertyCondition( AutomationElement.NameProperty , "Name" ) ) );

            var typeCondition = new AndCondition(
                new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.DataItem ) ,
                new OrCondition(
                    new PropertyCondition( AutomationElement.NameProperty , "Тип" ) ,
                    new PropertyCondition( AutomationElement.NameProperty , "Type" ) ) );

            foreach ( AutomationElement item in items )
            {
                var entry = new ChatEntry( );
                entry.Element = item;
                entry.OwnName = SafeCurrentName( item );

                AutomationElement nameItem = item.FindFirst( TreeScope.Descendants , nameCondition );
                if ( nameItem != null )
                {
                    entry.Name = GetLiveValue( nameItem );
                }

                AutomationElement typeItem = item.FindFirst( TreeScope.Descendants , typeCondition );
                if ( typeItem != null )
                {
                    entry.Type = GetLiveValue( typeItem );
                }

                entries.Add( entry );
            }

            return entries;
        }

        private static bool IsFolder( string type )
        {
            if ( string.IsNullOrEmpty( type ) )
            {
                return false;
            }
            return type.StartsWith( "Папка" , StringComparison.OrdinalIgnoreCase ) ||
                   type.StartsWith( "Folder" , StringComparison.OrdinalIgnoreCase );
        }

        private static string GetCachedString( AutomationElement element , AutomationProperty property )
        {
            try
            {
                object v = element.GetCachedPropertyValue( property );
                return v as string;
            }
            catch
            {
                return null;
            }
        }

        private static string GetCachedValue( AutomationElement element )
        {
            try
            {
                object v = element.GetCachedPropertyValue( ValuePattern.ValueProperty );
                return v as string;
            }
            catch
            {
                return null;
            }
        }

        private static string SafeCurrentName( AutomationElement element )
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

        private static string GetLiveValue( AutomationElement element )
        {
            try
            {
                object pattern;
                if ( element.TryGetCurrentPattern( ValuePattern.Pattern , out pattern ) )
                {
                    return ( (ValuePattern)pattern ).Current.Value;
                }
            }
            catch
            {
            }
            return null;
        }

        private static AutomationElement FindFirstChatList( int processId )
        {
            foreach ( AutomationElement window in TopWindows( processId ) )
            {
                AutomationElement chatList = window.FindFirst( TreeScope.Descendants ,
                    new AndCondition(
                        new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.List ) ,
                        new OrCondition(
                            new PropertyCondition( AutomationElement.NameProperty , "Чаты" ) ,
                            new PropertyCondition( AutomationElement.NameProperty , "Chats" ) ) ) );
                if ( chatList != null )
                {
                    return chatList;
                }
            }
            return null;
        }

        private static AutomationElementCollection TopWindows( int processId )
        {
            var condition = new PropertyCondition( AutomationElement.ProcessIdProperty , processId );
            return AutomationElement.RootElement.FindAll( TreeScope.Children , condition );
        }

        private static bool SelectChatItem( AutomationElement item )
        {
            try
            {
                object scroll;
                if ( item.TryGetCurrentPattern( ScrollItemPattern.Pattern , out scroll ) )
                {
                    ( (ScrollItemPattern)scroll ).ScrollIntoView( );
                    Thread.Sleep( 150 );
                }
            }
            catch
            {
            }

            if ( InvokeElement( item ) )
            {
                return true;
            }

            object pattern;
            if ( item.TryGetCurrentPattern( SelectionItemPattern.Pattern , out pattern ) )
            {
                try
                {
                    ( (SelectionItemPattern)pattern ).Select( );
                    return true;
                }
                catch
                {
                }
            }

            return ClickElement( item );
        }

        private static bool ClickElement( AutomationElement element )
        {
            try
            {
                System.Windows.Rect rect = element.Current.BoundingRectangle;
                if ( rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0 )
                {
                    return false;
                }

                int x = (int)( rect.Left + rect.Width / 2 );
                int y = (int)( rect.Top + rect.Height / 2 );

                SetCursorPos( x , y );
                Thread.Sleep( 50 );
                mouse_event( MOUSEEVENTF_LEFTDOWN , 0 , 0 , 0 , UIntPtr.Zero );
                mouse_event( MOUSEEVENTF_LEFTUP , 0 , 0 , 0 , UIntPtr.Zero );
                return true;
            }
            catch
            {
            }
            return false;
        }

        private static AutomationElement FindSendButton( AutomationElement root )
        {
            var condition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.Button );
            AutomationElementCollection buttons = root.FindAll( TreeScope.Descendants , condition );

            foreach ( AutomationElement button in buttons )
            {
                string name = ( button.Current.Name ?? string.Empty ).Trim( );
                if ( name.Length == 0 )
                {
                    continue;
                }

                string lower = name.ToLowerInvariant( );

                if ( lower.Contains( "сообщен" ) || lower.Contains( "message" ) )
                {
                    continue;
                }

                if ( lower.StartsWith( "отправ" ) || lower == "send" || lower.StartsWith( "send " ) )
                {
                    if ( button.Current.IsEnabled )
                    {
                        return button;
                    }
                }
            }

            return null;
        }

        private static void SetCaption( AutomationElement root , string caption )
        {
            try
            {
                var condition = new PropertyCondition( AutomationElement.ControlTypeProperty , ControlType.Edit );
                AutomationElementCollection edits = root.FindAll( TreeScope.Descendants , condition );

                foreach ( AutomationElement edit in edits )
                {
                    string name = edit.Current.Name ?? string.Empty;
                    if ( name.IndexOf( "Сообщение" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                         name.IndexOf( "Message" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                         name.IndexOf( "Поиск" , StringComparison.OrdinalIgnoreCase ) >= 0 ||
                         name.IndexOf( "Search" , StringComparison.OrdinalIgnoreCase ) >= 0 )
                    {
                        continue;
                    }

                    object pattern;
                    if ( edit.TryGetCurrentPattern( ValuePattern.Pattern , out pattern ) )
                    {
                        ( (ValuePattern)pattern ).SetValue( caption );
                        return;
                    }
                }
            }
            catch
            {
            }
        }

        private static bool InvokeElement( AutomationElement element )
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
            return false;
        }

        private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const int MOUSEEVENTF_LEFTUP = 0x0004;

        [DllImport( "user32.dll" )]
        private static extern bool SetCursorPos( int x , int y );

        [DllImport( "user32.dll" )]
        private static extern void mouse_event( int dwFlags , int dx , int dy , int dwData , UIntPtr dwExtraInfo );
    }
}
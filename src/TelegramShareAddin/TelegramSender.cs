using System;
using System.Collections.Generic;
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
        Native = 1
    }

    /// <summary>Мессенджер, в который отправляется файл.</summary>
    public enum Messenger
    {
        Telegram = 0,
        Max = 1
    }

    /// <summary>Параметры отправки.</summary>
    public class SendOptions
    {
        public string Contact { get; set; }
        public string Message { get; set; }

        /// <summary>true — контакт выбран из списка чатов (открываем чат кликом).</summary>
        public bool OpenByChatName { get; set; }
    }

    /// <summary>Информация о чате мессенджера.</summary>
    public class MessengerChatInfo
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

        private readonly Settings _settings;

        private static string _cachedExePath;

        public TelegramSender( Settings settings )
        {
            _settings = settings ?? new Settings( );
        }

        /// <summary>Путь во временной папке с заданным именем файла (имя делается уникальным).</summary>
        public static string NewTempFile( string fileName )
        {
            string dir = Path.Combine( Path.GetTempPath( ) , "TelegramShare" );
            Directory.CreateDirectory( dir );

            string name = SanitizeFileName( fileName );
            string ext = Path.GetExtension( name );
            string baseName = Path.GetFileNameWithoutExtension( name );

            // Уникальный суффикс: две одновременные отправки не перетирают файл друг друга.
            string unique = baseName + "_" + Guid.NewGuid( ).ToString( "N" ).Substring( 0 , 8 ) + ext;

            return Path.Combine( dir , unique );
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

        /// <summary>Удалить временные файлы прошлых отправок (старше суток) только из своей папки.</summary>
        public static void CleanupOldTempFiles( )
        {
            DeleteOld( Path.Combine( Path.GetTempPath( ) , "TelegramShare" ) );
        }

        private static void DeleteOld( string dir )
        {
            try
            {
                if ( !Directory.Exists( dir ) )
                {
                    return;
                }
                foreach ( string file in Directory.GetFiles( dir ) )
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
        public void SendPreparedFile( SendOptions options , string tempFile )
        {
            if ( string.IsNullOrWhiteSpace( options.Contact ) )
            {
                throw new ArgumentException( "Не указан контакт." , "Contact" );
            }

            try
            {
                ClipboardHelper.SetFileAndPath( tempFile );
                OpenChatAndSend( options );
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

        /// <summary>Открыть в Telegram окно «Переслать…» для уже подготовленного файла.</summary>
        public void SendWithPickerPreparedFile( string tempFile )
        {
            SendWithPickerPreparedFiles( new string[ ] { tempFile } );
        }

        /// <summary>Открыть в Telegram окно «Переслать…» сразу для нескольких файлов.</summary>
        public void SendWithPickerPreparedFiles( IEnumerable<string> files )
        {
            var list = new List<string>( files );
            if ( list.Count == 0 )
            {
                throw new InvalidOperationException( "Нет файлов для отправки." );
            }

            LaunchTelegramSendPath( list );

            // После сохранения полный путь к файлу остаётся в буфере обмена.
            try
            {
                ClipboardHelper.SetText( string.Join( Environment.NewLine , list ) );
            }
            catch
            {
            }
        }

        private void OpenChatAndSend( SendOptions options )
        {
            string contact = options.Contact.Trim( );
            bool opened;

            if ( options.OpenByChatName )
            {
                int telegramPid = GetTelegramProcessId( );
                if ( telegramPid == 0 )
                {
                    throw new InvalidOperationException( "Telegram Desktop не запущен. Запустите Telegram и повторите." );
                }

                IntPtr hwnd = FindTelegramWindow( );
                if ( hwnd == IntPtr.Zero )
                {
                    throw new InvalidOperationException( "Не найдено окно Telegram Desktop. Разверните окно Telegram из трея." );
                }

                UiWin.EnsureForeground( hwnd );
                Thread.Sleep( 300 );

                opened = TelegramUi.OpenChat( telegramPid , contact , 8000 );

                if ( !opened )
                {
                    // Чат мог оказаться ниже в списке: пробуем открыть по ссылке.
                    AddinCore.Log( "Send: чат не найден в списке, пробую ссылку" );
                    opened = TryOpenByUri( contact );
                }

                if ( !opened )
                {
                    throw new InvalidOperationException( "Чат «" + contact + "» не найден в списке Telegram. " +
                        "Откройте его в самом Telegram и повторите отправку — либо выберите чат кнопкой «Обновить из Telegram»." );
                }

                Thread.Sleep( _settings.ChatOpenDelay );
            }
            else
            {
                LaunchTelegram( BuildTelegramUri( contact ) );
                Thread.Sleep( _settings.ChatOpenDelay );
            }

            IntPtr window = FindTelegramWindow( );
            if ( window == IntPtr.Zero )
            {
                throw new InvalidOperationException( "Не найдено окно Telegram Desktop. Разверните окно Telegram из трея и повторите." );
            }

            if ( !UiWin.EnsureForeground( window ) )
            {
                throw new InvalidOperationException( "Не удалось вывести Telegram на передний план — отправка отменена." );
            }

            int pid = UiWin.ProcessIdOf( window );

            // Клавиши уходят в активное окно: вставляем только когда фокус точно в Telegram.
            UiWin.SendKeysTo( pid , "^v" , "Telegram" );
            AddinCore.Log( "Send: файл вставлен, ожидаю диалог Telegram" );
            Thread.Sleep( _settings.PasteDelay );

            bool confirmed = TelegramUi.ConfirmSend( options.Message , pid , 8000 );
            AddinCore.Log( "Send: UIA подтверждение = " + confirmed );

            if ( !confirmed )
            {
                // Резервный путь — тоже только при фокусе в Telegram, иначе текст уйдёт чужому приложению.
                if ( !string.IsNullOrEmpty( options.Message ) )
                {
                    UiWin.SendKeysTo( pid , EscapeSendKeys( options.Message ) , "Telegram" );
                    Thread.Sleep( 300 );
                }

                UiWin.SendKeysTo( pid , "{ENTER}" , "Telegram" );
            }
        }

        private static bool TryOpenByUri( string contact )
        {
            string c = ( contact ?? string.Empty ).Trim( );

            // По ссылке открываем только то, что действительно является username,
            // телефоном или числовым ID. Иначе Telegram показал бы ошибку, а вставка
            // ушла бы в случайно открытый чат — это недопустимо.
            bool username = c.StartsWith( "@" , StringComparison.Ordinal ) && c.Length > 1;
            bool phone = c.StartsWith( "+" , StringComparison.Ordinal ) && c.Length > 4;
            bool numericId = c.Length > 0 && c.All( char.IsDigit );

            if ( !username && !phone && !numericId )
            {
                AddinLog.Warn( "Send: «" + c + "» не похоже на username/телефон/ID — ссылку не открываю" );
                return false;
            }

            try
            {
                LaunchTelegram( BuildTelegramUri( c ) );
                return true;
            }
            catch ( Exception ex )
            {
                AddinLog.Warn( "Send: ссылка не открылась — " + AddinLog.Describe( ex ) );
                return false;
            }
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

        private static void LaunchTelegramSendPath( IList<string> files )
        {
            string exe = FindTelegramExe( );
            if ( string.IsNullOrEmpty( exe ) )
            {
                throw new InvalidOperationException( "Не найден Telegram Desktop (Telegram.exe). Установите Telegram или запустите его один раз." );
            }

            var arguments = new StringBuilder( "-sendpath" );
            foreach ( string file in files )
            {
                arguments.Append( " \"" ).Append( file ).Append( '"' );
            }

            try
            {
                Process.Start( new ProcessStartInfo( exe , arguments.ToString( ) ) { UseShellExecute = false } );
            }
            catch ( Exception ex )
            {
                throw new InvalidOperationException( "Не удалось открыть Telegram с файлом: " + ex.Message , ex );
            }
        }

        private static string FindTelegramExe( )
        {
            if ( !string.IsNullOrEmpty( _cachedExePath ) && File.Exists( _cachedExePath ) )
            {
                return _cachedExePath;
            }

            Process[ ] processes = UiWin.FindProcesses( "Telegram" );
            try
            {
                foreach ( Process p in processes )
                {
                    try
                    {
                        string path = p.MainModule.FileName;
                        if ( !string.IsNullOrEmpty( path ) && File.Exists( path ) )
                        {
                            _cachedExePath = path;
                            return path;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            finally
            {
                foreach ( Process p in processes )
                {
                    p.Dispose( );
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
            Process[ ] processes = UiWin.FindProcesses( "Telegram" );
            try
            {
                foreach ( Process p in processes )
                {
                    try
                    {
                        IntPtr h = p.MainWindowHandle;
                        if ( h != IntPtr.Zero && UiWin.IsWindowVisible( h ) )
                        {
                            return h;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            finally
            {
                foreach ( Process p in processes )
                {
                    p.Dispose( );
                }
            }

            return IntPtr.Zero;
        }

        public static int GetTelegramProcessId( )
        {
            Process[ ] processes = UiWin.FindProcesses( "Telegram" );
            try
            {
                foreach ( Process p in processes )
                {
                    try
                    {
                        if ( p.MainWindowHandle != IntPtr.Zero )
                        {
                            return p.Id;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            finally
            {
                foreach ( Process p in processes )
                {
                    p.Dispose( );
                }
            }

            return 0;
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

        public static List<MessengerChatInfo> GetChats( int processId )
        {
            var result = new List<MessengerChatInfo>( );
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
                    result.Add( new MessengerChatInfo { Name = e.Name , Type = e.Type } );
                }
            }

            // «Избранное» (Saved Messages) — особая строка списка, у неё нет поля
            // «Название». Добавляем её отдельно, иначе отправить туда нельзя.
            if ( !result.Exists( c => IsSavedMessages( c.Name ) ) )
            {
                foreach ( ChatEntry e in entries )
                {
                    string saved = SavedMessagesName( e.OwnName );
                    if ( saved != null && seen.Add( saved ) )
                    {
                        result.Insert( 0 , new MessengerChatInfo { Name = saved , Type = string.Empty } );
                        break;
                    }
                }
            }

            return result;
        }

        /// <summary>Имя первого сегмента, если элемент списка — «Избранное»; иначе null.</summary>
        private static string SavedMessagesName( string ownName )
        {
            if ( string.IsNullOrEmpty( ownName ) )
            {
                return null;
            }

            string first = ownName.Split( ',' )[ 0 ].Trim( );
            return IsSavedMessages( first ) ? first : null;
        }

        private static bool IsSavedMessages( string name )
        {
            if ( string.IsNullOrEmpty( name ) )
            {
                return false;
            }

            string n = name.Trim( );

            return n.StartsWith( "Избранное" , StringComparison.OrdinalIgnoreCase ) ||
                   n.StartsWith( "Избранные" , StringComparison.OrdinalIgnoreCase ) ||
                   n.StartsWith( "Saved Messages" , StringComparison.OrdinalIgnoreCase ) ||
                   n.StartsWith( "Saved messages" , StringComparison.OrdinalIgnoreCase );
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
                AutomationElement chatList;
                try
                {
                    chatList = FindFirstChatList( processId );
                }
                catch
                {
                    chatList = null;
                }

                if ( chatList == null )
                {
                    Thread.Sleep( 300 );
                    continue;
                }

                if ( SelectByName( chatList , chatName ) )
                {
                    return true;
                }

                // Чат мог быть ниже видимой части списка — листаем вниз.
                if ( !ScrollList( chatList , ScrollAmount.LargeIncrement ) )
                {
                    break;
                }

                Thread.Sleep( 250 );
            }

            return false;
        }

        private static bool SelectByName( AutomationElement chatList , string chatName )
        {
            List<ChatEntry> entries;
            try
            {
                entries = ReadChatEntries( chatList );
            }
            catch
            {
                return false;
            }

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

            // «Избранное»: имя лежит не в поле «Название», а в строке элемента.
            if ( IsSavedMessages( chatName ) )
            {
                foreach ( ChatEntry e in entries )
                {
                    if ( SavedMessagesName( e.OwnName ) != null && SelectChatItem( e.Element ) )
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ScrollList( AutomationElement list , ScrollAmount amount )
        {
            try
            {
                object pattern;
                if ( !list.TryGetCurrentPattern( ScrollPattern.Pattern , out pattern ) )
                {
                    return false;
                }

                var scroll = (ScrollPattern) pattern;
                if ( !scroll.Current.VerticallyScrollable || scroll.Current.VerticalScrollPercent >= 99 )
                {
                    return false;
                }

                scroll.Scroll( ScrollAmount.NoAmount , amount );
                return true;
            }
            catch
            {
                return false;
            }
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

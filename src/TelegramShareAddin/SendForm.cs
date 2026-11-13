using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Диалог отправки: мессенджер (Telegram/MAX), список чатов с фильтром,
    /// выбор мышью, поле контакта, сообщение и формат файла.
    /// Групповые чаты Telegram отображаются курсивом.
    /// </summary>
    public sealed class SendForm : Form
    {
        private sealed class ContactItem
        {
            public string Name;
            public bool Group;

            public override string ToString( )
            {
                return Name;
            }
        }

        private readonly Settings _settings;
        private readonly string _nativeLabel;
        private readonly bool _allowPdf;
        private readonly List<ContactItem> _all = new List<ContactItem>( );
        private readonly HashSet<string> _chatNames = new HashSet<string>( StringComparer.OrdinalIgnoreCase );
        private readonly HashSet<string> _hiddenChats = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

        private ComboBox _messengerBox;
        private TextBox _searchBox;
        private ListBox _contactsList;
        private TextBox _contactBox;
        private TextBox _messageBox;
        private ComboBox _formatBox;
        private CheckBox _rememberCheckBox;
        private Label _statusLabel;
        private Label _contactsLabel;
        private Label _contactLabel;
        private Label _hintLabel;
        private Button _refreshButton;
        private Font _italicFont;

        public string Contact
        {
            get { return _contactBox.Text.Trim( ); }
        }

        public string Message
        {
            get { return _messageBox.Text; }
        }

        public Messenger Messenger
        {
            get { return _messengerBox.SelectedIndex == 1 ? Messenger.Max : Messenger.Telegram; }
        }

        public SendFileFormat Format
        {
            get
            {
                if ( !_allowPdf )
                {
                    return SendFileFormat.Docx;
                }
                return _formatBox.SelectedIndex == 1 ? SendFileFormat.Docx : SendFileFormat.Pdf;
            }
        }

        public bool RememberContact
        {
            get { return _rememberCheckBox.Checked; }
        }

        /// <summary>true — контакт совпадает с чатом из загруженного списка (открываем чат кликом(.</summary>
        public bool OpenByChatName
        {
            get { return _chatNames.Contains( Contact ); }
        }

        public SendForm( Settings settings , string nativeLabel , bool allowPdf , Messenger defaultMessenger )
        {
            _settings = settings ?? new Settings( );
            _nativeLabel = string.IsNullOrEmpty( nativeLabel ) ? "Копия" : nativeLabel;
            _allowPdf = allowPdf;

            Text = "Отправить сообщение";
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size( 560 , 572 );
            MinimumSize = new Size( 576 , 620 );
            Font = new Font( "Segoe UI" , 9F );
            AutoScaleMode = AutoScaleMode.None;

            _italicFont = new Font( Font.FontFamily , Font.Size , FontStyle.Italic );

            InitializeControls( defaultMessenger );

            RebuildAll( );
            StartChatLoad( );
        }

        private void InitializeControls( Messenger defaultMessenger )
        {
            int y = 12;

            var messengerLabel = new Label
            {
                Left = 12,
                Top = y + 3,
                Width = 92,
                Height = 18,
                Text = "Мессенджер:"
            };
            Controls.Add( messengerLabel );

            _messengerBox = new ComboBox
            {
                Left = 108,
                Top = y,
                Width = 220,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _messengerBox.Items.AddRange( new object[ ] { "Telegram" , "MAX" } );
            _messengerBox.SelectedIndex = defaultMessenger == Messenger.Max ? 1 : 0;
            _messengerBox.SelectedIndexChanged += OnMessengerChanged;
            Controls.Add( _messengerBox );
            y += 36;

            _contactsLabel = new Label
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 18,
                Text = "Контакты — выберите кликом; курсив — групповые чаты:"
            };
            Controls.Add( _contactsLabel );
            y += 20;

            _searchBox = new TextBox
            {
                Left = 12,
                Top = y,
                Width = 536
            };
            _searchBox.TextChanged += OnSearchChanged;
            Controls.Add( _searchBox );
            y += 26;

            _contactsList = new ListBox
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 110,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 18
            };
            _contactsList.SelectedIndexChanged += OnContactSelected;
            _contactsList.DrawItem += OnDrawItem;
            Controls.Add( _contactsList );
            y += 116;

            var removeButton = new Button
            {
                Left = 12,
                Top = y,
                Width = 170,
                Height = 26,
                Text = "Удалить из списка"
            };
            removeButton.Click += OnRemoveContact;
            Controls.Add( removeButton );

            _refreshButton = new Button
            {
                Left = 190,
                Top = y,
                Width = 170,
                Height = 26,
                Text = "Обновить из Telegram"
            };
            _refreshButton.Click += OnRefreshChats;
            Controls.Add( _refreshButton );

            _statusLabel = new Label
            {
                Left = 368,
                Top = y + 5,
                Width = 180,
                Height = 18,
                ForeColor = SystemColors.GrayText,
                Text = ""
            };
            Controls.Add( _statusLabel );
            y += 40;

            _contactLabel = new Label
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 18,
                Text = "Контакт — username, +номер телефона или числовой ID:"
            };
            Controls.Add( _contactLabel );
            y += 20;

            _contactBox = new TextBox
            {
                Left = 12,
                Top = y,
                Width = 392,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.CustomSource
            };
            Controls.Add( _contactBox );

            var addButton = new Button
            {
                Left = 412,
                Top = y - 1,
                Width = 136,
                Height = 26,
                Text = "В список"
            };
            addButton.Click += OnAddContact;
            Controls.Add( addButton );
            y += 30;

            _hintLabel = new Label
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 16,
                ForeColor = SystemColors.GrayText,
                Text = "Примеры: @durov, +79123456789 или 123456789"
            };
            _hintLabel.Font = new Font( "Segoe UI" , 8F );
            Controls.Add( _hintLabel );
            y += 26;

            AddLabel( "Сообщение — необязательно:" , 12 , y );
            y += 20;

            _messageBox = new TextBox
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 60,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                AcceptsReturn = false
            };
            Controls.Add( _messageBox );
            y += 72;

            AddLabel( "Формат файла:" , 12 , y );
            y += 20;

            _formatBox = new ComboBox
            {
                Left = 12,
                Top = y,
                Width = 536,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            if ( _allowPdf )
            {
                _formatBox.Items.AddRange( new object[ ] { "PDF (рекомендуется)" , _nativeLabel } );
                _formatBox.SelectedIndex = _settings.LastFormat == "DOCX" ? 1 : 0;
            }
            else
            {
                _formatBox.Items.Add( _nativeLabel );
                _formatBox.SelectedIndex = 0;
            }
            Controls.Add( _formatBox );
            y += 34;

            _rememberCheckBox = new CheckBox
            {
                Left = 12,
                Top = y,
                Width = 536,
                Text = "Запоминать последний контакт и формат",
                Checked = _settings.RememberContact
            };
            Controls.Add( _rememberCheckBox );
            y += 30;

            var okButton = new Button
            {
                Left = 12,
                Top = y,
                Width = 150,
                Height = 34,
                Text = "Отправить"
            };
            okButton.Click += OnOkClicked;

            var cancelButton = new Button
            {
                Left = 398,
                Top = y,
                Width = 150,
                Height = 34,
                Text = "Отмена",
                DialogResult = DialogResult.Cancel
            };

            Controls.Add( okButton );
            Controls.Add( cancelButton );

            AcceptButton = okButton;
            CancelButton = cancelButton;

            UpdateMessengerTexts( );
        }

        private const int EM_SETCUEBANNER = 0x1501;

        [DllImport( "user32.dll" , CharSet = CharSet.Unicode )]
        private static extern IntPtr SendMessage( IntPtr hWnd , int msg , IntPtr wParam , string lParam );

        protected override void OnLoad( EventArgs e )
        {
            base.OnLoad( e );
            try
            {
                SendMessage( _searchBox.Handle , EM_SETCUEBANNER , (IntPtr)1 , "Фильтр по первым буквам..." );
            }
            catch
            {
            }
        }

        // --------------------------------------------------------------
        // Мессенджер
        // --------------------------------------------------------------

        private void OnMessengerChanged( object sender , EventArgs e )
        {
            _chatInfos.Clear( );
            _chatNames.Clear( );
            UpdateMessengerTexts( );
            RebuildAll( );
            StartChatLoad( );
        }

        private void UpdateMessengerTexts( )
        {
            bool max = Messenger == Messenger.Max;

            _contactsLabel.Text = max
                ? "Чаты MAX — выберите кликом:"
                : "Контакты — выберите кликом; курсив — групповые чаты:";

            _refreshButton.Text = max ? "Обновить из MAX" : "Обновить из Telegram";

            _contactLabel.Text = max
                ? "Контакт — имя чата из списка или @username:"
                : "Контакт — username, +номер телефона или числовой ID:";

            _hintLabel.Text = max
                ? "Примеры: Иван Петров или @support"
                : "Примеры: @durov, +79123456789 или 123456789";
        }

        // --------------------------------------------------------------
        // Загрузка чатов
        // --------------------------------------------------------------

        private void StartChatLoad( )
        {
            Messenger messenger = Messenger;

            _statusLabel.Text = messenger == Messenger.Max
                ? "Загрузка из MAX..."
                : "Загрузка из Telegram...";

            var thread = new Thread( new ThreadStart( delegate { LoadChatsWorker( messenger ); } ) );
            thread.IsBackground = true;
            thread.SetApartmentState( ApartmentState.STA );
            thread.Start( );
        }

        private void LoadChatsWorker( Messenger messenger )
        {
            var chats = new List<MessengerChatInfo>( );
            bool running = false;
            try
            {
                int pid = messenger == Messenger.Max
                    ? MaxSender.GetProcessId( )
                    : TelegramSender.GetTelegramProcessId( );

                if ( pid != 0 )
                {
                    running = true;
                    chats = messenger == Messenger.Max
                        ? MaxUi.GetChats( pid )
                        : TelegramUi.GetChats( pid );
                }
            }
            catch
            {
            }

            try
            {
                if ( !IsDisposed )
                {
                    BeginInvoke( (Action)delegate { ApplyChats( chats , running , messenger ); } );
                }
            }
            catch
            {
            }
        }

        private void ApplyChats( List<MessengerChatInfo> chats , bool running , Messenger messenger )
        {
            if ( messenger != Messenger )
            {
                return;
            }

            _chatInfos = chats;
            _chatNames.Clear( );

            foreach ( MessengerChatInfo c in chats )
            {
                if ( !string.IsNullOrEmpty( c.Name ) )
                {
                    _chatNames.Add( c.Name );
                }
            }

            string name = messenger == Messenger.Max ? "MAX" : "Telegram";

            if ( chats.Count > 0 )
            {
                _statusLabel.Text = "Из " + name + ": " + chats.Count;
            }
            else if ( !running )
            {
                _statusLabel.Text = name + " не запущен";
            }
            else
            {
                _statusLabel.Text = "Список из " + name + " пуст";
            }

            RebuildAll( );
        }

        // --------------------------------------------------------------
        // Список контактов, фильтр, отрисовка
        // --------------------------------------------------------------

        private void RebuildAll( )
        {
            string keep = Contact;

            _all.Clear( );
            var seen = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

            foreach ( string c in _settings.RecentContacts )
            {
                if ( seen.Add( c ) )
                {
                    _all.Add( new ContactItem { Name = c , Group = false } );
                }
            }

            foreach ( string c in _chatNames )
            {
                if ( _hiddenChats.Contains( c ) )
                {
                    continue;
                }

                bool isGroup = false;
                MessengerChatInfo info = _chatInfo( c );
                if ( info != null )
                {
                    isGroup = info.IsGroup;
                }

                if ( seen.Add( c ) )
                {
                    _all.Add( new ContactItem { Name = c , Group = isGroup } );
                }
            }

            _contactBox.AutoCompleteCustomSource = new AutoCompleteStringCollection( );
            int limit = Math.Min( _all.Count , 200 );
            for ( int i = 0; i < limit; i++ )
            {
                _contactBox.AutoCompleteCustomSource.Add( _all[i].Name );
            }

            ApplyFilter( keep );
        }

        private List<MessengerChatInfo> _chatInfos = new List<MessengerChatInfo>( );

        private MessengerChatInfo _chatInfo( string name )
        {
            foreach ( MessengerChatInfo c in _chatInfos )
            {
                if ( string.Equals( c.Name , name , StringComparison.OrdinalIgnoreCase ) )
                {
                    return c;
                }
            }
            return null;
        }

        private void ApplyFilter( string keep )
        {
            string filter = _searchBox.Text.Trim( );

            _contactsList.BeginUpdate( );
            _contactsList.Items.Clear( );

            foreach ( ContactItem item in _all )
            {
                if ( filter.Length == 0 ||
                     item.Name.StartsWith( filter , StringComparison.OrdinalIgnoreCase ) )
                {
                    _contactsList.Items.Add( item );
                }
            }

            _contactsList.EndUpdate( );

            if ( !string.IsNullOrEmpty( keep ) )
            {
                int index = FindIndex( keep );
                if ( index >= 0 )
                {
                    _contactsList.SelectedIndex = index;
                    _contactBox.Text = keep;
                    return;
                }
            }

            if ( _contactsList.Items.Count > 0 && _contactsList.SelectedIndex < 0 )
            {
                _contactsList.SelectedIndex = 0;
            }
        }

        private int FindIndex( string name )
        {
            for ( int i = 0; i < _contactsList.Items.Count; i++ )
            {
                ContactItem item = _contactsList.Items[i] as ContactItem;
                if ( item != null && string.Equals( item.Name , name , StringComparison.OrdinalIgnoreCase ) )
                {
                    return i;
                }
            }
            return -1;
        }

        private void OnSearchChanged( object sender , EventArgs e )
        {
            ApplyFilter( null );
        }

        private void OnDrawItem( object sender , DrawItemEventArgs e )
        {
            if ( e.Index < 0 )
            {
                return;
            }

            e.DrawBackground( );

            ContactItem item = _contactsList.Items[e.Index] as ContactItem;
            if ( item != null )
            {
                Font font = item.Group ? _italicFont : _contactsList.Font;
                Brush brush = ( ( e.State & DrawItemState.Selected ) == DrawItemState.Selected )
                    ? SystemBrushes.HighlightText
                    : SystemBrushes.ControlText;
                e.Graphics.DrawString( item.Name , font , brush , e.Bounds );
            }

            e.DrawFocusRectangle( );
        }

        private void OnContactSelected( object sender , EventArgs e )
        {
            ContactItem item = _contactsList.SelectedItem as ContactItem;
            if ( item != null )
            {
                _contactBox.Text = item.Name;
            }
        }

        private void OnAddContact( object sender , EventArgs e )
        {
            string c = _contactBox.Text.Trim( ).TrimStart( '@' );
            if ( string.IsNullOrEmpty( c ) )
            {
                MessageBox.Show( "Введите контакт, чтобы добавить его в список." , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                return;
            }

            _hiddenChats.Remove( c );
            if ( !_settings.RecentContacts.Contains( c ) )
            {
                _settings.RecentContacts.Insert( 0 , c );
            }

            RebuildAll( );
            _contactBox.Text = c;
        }

        private void OnRemoveContact( object sender , EventArgs e )
        {
            ContactItem item = _contactsList.SelectedItem as ContactItem;
            if ( item == null )
            {
                return;
            }

            if ( _settings.RecentContacts.Remove( item.Name ) )
            {
                RebuildAll( );
                return;
            }

            _hiddenChats.Add( item.Name );
            RebuildAll( );
        }

        private void OnRefreshChats( object sender , EventArgs e )
        {
            StartChatLoad( );
        }

        private void AddLabel( string text , int left , int top )
        {
            var label = new Label
            {
                Left = left,
                Top = top,
                Width = 536,
                Height = 18,
                Text = text
            };
            Controls.Add( label );
        }

        private void OnOkClicked( object sender , EventArgs e )
        {
            if ( string.IsNullOrWhiteSpace( Contact ) )
            {
                MessageBox.Show( "Укажите контакт: выберите из списка или введите вручную." , "Отправить сообщение" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                return;
            }

            DialogResult = DialogResult.OK;
            Close( );
        }
    }
}

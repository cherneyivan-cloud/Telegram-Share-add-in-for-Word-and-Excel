using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Диалог отправки: список контактов (сохранённые + из Telegram) с фильтром,
    /// выбор мышью, поле контакта, сообщение и формат файла.
    /// Групповые чаты и каналы отображаются курсивом.
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
        private readonly List<ContactItem> _all = new List<ContactItem>( );
        private readonly HashSet<string> _telegramChats = new HashSet<string>( StringComparer.OrdinalIgnoreCase );
        private readonly HashSet<string> _hiddenChats = new HashSet<string>( StringComparer.OrdinalIgnoreCase );

        private TextBox _searchBox;
        private ListBox _contactsList;
        private TextBox _contactBox;
        private TextBox _messageBox;
        private ComboBox _formatBox;
        private CheckBox _rememberCheckBox;
        private Label _statusLabel;
        private Font _italicFont;

        public string Contact
        {
            get { return _contactBox.Text.Trim( ); }
        }

        public string Message
        {
            get { return _messageBox.Text; }
        }

        public SendFileFormat Format
        {
            get { return _formatBox.SelectedIndex == 1 ? SendFileFormat.Docx : SendFileFormat.Pdf; }
        }

        public bool RememberContact
        {
            get { return _rememberCheckBox.Checked; }
        }

        /// <summary>true — контакт совпадает с чатом из списка Telegram (открываем чат кликом(.</summary>
        public bool OpenByChatName
        {
            get { return _telegramChats.Contains( Contact ); }
        }

        public SendForm( Settings settings , string nativeLabel )
        {
            _settings = settings ?? new Settings( );
            _nativeLabel = string.IsNullOrEmpty( nativeLabel ) ? "Копия" : nativeLabel;

            Text = "Отправить в Telegram";
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size( 560 , 520 );
            MinimumSize = new Size( 576 , 560 );
            Font = new Font( "Segoe UI" , 9F );
            AutoScaleMode = AutoScaleMode.None;

            _italicFont = new Font( Font.FontFamily , Font.Size , FontStyle.Italic );

            InitializeControls( );

            RebuildAll( );
            StartTelegramLoad( );
        }

        private void InitializeControls( )
        {
            int y = 12;

            AddLabel( "Контакты — выберите кликом; курсив — групповые чаты:" , 12 , y );
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

            var refreshButton = new Button
            {
                Left = 190,
                Top = y,
                Width = 170,
                Height = 26,
                Text = "Обновить из Telegram"
            };
            refreshButton.Click += OnRefreshTelegram;
            Controls.Add( refreshButton );

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

            AddLabel( "Контакт — username, +номер телефона или числовой ID:" , 12 , y );
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

            var hint = new Label
            {
                Left = 12,
                Top = y,
                Width = 536,
                Height = 16,
                ForeColor = SystemColors.GrayText,
                Text = "Примеры: @durov, +79123456789 или 123456789"
            };
            hint.Font = new Font( "Segoe UI" , 8F );
            Controls.Add( hint );
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
            _formatBox.Items.AddRange( new object[ ] { "PDF (рекомендуется)" , _nativeLabel } );
            _formatBox.SelectedIndex = _settings.LastFormat == "DOCX" ? 1 : 0;
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
        // Загрузка чатов из Telegram
        // --------------------------------------------------------------

        private void StartTelegramLoad( )
        {
            _statusLabel.Text = "Загрузка из Telegram...";

            var thread = new Thread( LoadTelegramChatsWorker );
            thread.IsBackground = true;
            thread.SetApartmentState( ApartmentState.STA );
            thread.Start( );
        }

        private void LoadTelegramChatsWorker( )
        {
            var chats = new List<TelegramChatInfo>( );
            bool running = false;
            try
            {
                int pid = TelegramSender.GetTelegramProcessId( );
                if ( pid != 0 )
                {
                    running = true;
                    chats = TelegramUi.GetChats( pid );
                }
            }
            catch
            {
            }

            try
            {
                if ( !IsDisposed )
                {
                    BeginInvoke( (Action)delegate { ApplyTelegramChats( chats , running ); } );
                }
            }
            catch
            {
            }
        }

        private void ApplyTelegramChats( List<TelegramChatInfo> chats , bool running )
        {
            _chatInfos = chats;
            _telegramChats.Clear( );

            foreach ( TelegramChatInfo c in chats )
            {
                if ( !string.IsNullOrEmpty( c.Name ) )
                {
                    _telegramChats.Add( c.Name );
                }
            }

            if ( chats.Count > 0 )
            {
                _statusLabel.Text = "Из Telegram: " + chats.Count;
            }
            else if ( !running )
            {
                _statusLabel.Text = "Telegram не запущен";
            }
            else
            {
                _statusLabel.Text = "Список из Telegram пуст";
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

            foreach ( string c in _telegramChats )
            {
                if ( _hiddenChats.Contains( c ) )
                {
                    continue;
                }

                bool isGroup = false;
                TelegramChatInfo info = _chatInfo( c );
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

        private List<TelegramChatInfo> _chatInfos = new List<TelegramChatInfo>( );

        private TelegramChatInfo _chatInfo( string name )
        {
            foreach ( TelegramChatInfo c in _chatInfos )
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
                MessageBox.Show( "Введите контакт, чтобы добавить его в список." , "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Information );
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

        private void OnRefreshTelegram( object sender , EventArgs e )
        {
            StartTelegramLoad( );
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
                MessageBox.Show( "Укажите контакт: выберите из списка или введите вручную." , "Telegram Share" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                return;
            }

            DialogResult = DialogResult.OK;
            Close( );
        }
    }
}
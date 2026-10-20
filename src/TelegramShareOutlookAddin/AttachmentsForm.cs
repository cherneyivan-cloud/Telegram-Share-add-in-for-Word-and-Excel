using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>Выбор отдельных вложений текущего письма Outlook.</summary>
    public sealed class AttachmentsForm : Form
    {
        public sealed class Item
        {
            public int Index;
            public string Name;
            public long Size;

            public override string ToString( )
            {
                return Name + "   (" + ( Size / 1024 ) + " КБ)";
            }
        }

        private readonly CheckedListBox _list;

        public List<int> SelectedIndices
        {
            get
            {
                var result = new List<int>( );
                foreach ( object o in _list.CheckedItems )
                {
                    Item item = o as Item;
                    if ( item != null )
                    {
                        result.Add( item.Index );
                    }
                }
                return result;
            }
        }

        public AttachmentsForm( List<Item> items )
        {
            Text = "Вложения письма";
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size( 460 , 380 );
            MinimumSize = new Size( 430 , 330 );
            Font = new Font( "Segoe UI" , 9F );
            AutoScaleMode = AutoScaleMode.None;

            var label = new Label
            {
                Left = 12,
                Top = 12,
                Width = 436,
                Height = 18,
                Text = "Выберите вложения для отправки в Telegram:"
            };
            Controls.Add( label );

            _list = new CheckedListBox
            {
                Left = 12,
                Top = 36,
                Width = 436,
                Height = 260,
                CheckOnClick = true,
                IntegralHeight = false
            };
            foreach ( Item item in items )
            {
                _list.Items.Add( item );
            }
            Controls.Add( _list );

            var all = new Button
            {
                Left = 12,
                Top = 308,
                Width = 120,
                Height = 30,
                Text = "Выбрать все"
            };
            all.Click += OnSelectAll;
            Controls.Add( all );

            var ok = new Button
            {
                Left = 252,
                Top = 308,
                Width = 100,
                Height = 30,
                Text = "Отправить"
            };
            ok.Click += OnOk;

            var cancel = new Button
            {
                Left = 360,
                Top = 308,
                Width = 88,
                Height = 30,
                Text = "Отмена",
                DialogResult = DialogResult.Cancel
            };

            Controls.Add( ok );
            Controls.Add( cancel );

            AcceptButton = ok;
            CancelButton = cancel;
        }

        private void OnSelectAll( object sender , EventArgs e )
        {
            for ( int i = 0; i < _list.Items.Count; i++ )
            {
                _list.SetItemChecked( i , true );
            }
        }

        private void OnOk( object sender , EventArgs e )
        {
            if ( SelectedIndices.Count == 0 )
            {
                MessageBox.Show( "Отметьте хотя бы одно вложение." , "Отправить в Telegram" , MessageBoxButtons.OK , MessageBoxIcon.Information );
                return;
            }

            DialogResult = DialogResult.OK;
            Close( );
        }
    }
}
using System;
using System.Drawing;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>Небольшой диалог выбора формата отправляемого файла.</summary>
    public sealed class FormatForm : Form
    {
        private RadioButton _pdf;
        private RadioButton _docx;

        public SendFileFormat Format
        {
            get { return _docx.Checked ? SendFileFormat.Docx : SendFileFormat.Pdf; }
        }

        public FormatForm( SendFileFormat current , string nativeLabel )
        {
            Text = "Формат файла";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size( 320 , 160 );
            Font = new Font( "Segoe UI" , 9F );
            AutoScaleMode = AutoScaleMode.None;

            var label = new Label
            {
                Left = 12,
                Top = 14,
                Width = 296,
                Height = 18,
                Text = "Выберите формат отправляемого файла:"
            };
            Controls.Add( label );

            _pdf = new RadioButton
            {
                Left = 20,
                Top = 46,
                Width = 280,
                Height = 22,
                Text = "PDF",
                Checked = current != SendFileFormat.Docx
            };
            _docx = new RadioButton
            {
                Left = 20,
                Top = 74,
                Width = 280,
                Height = 22,
                Text = nativeLabel,
                Checked = current == SendFileFormat.Docx
            };
            Controls.Add( _pdf );
            Controls.Add( _docx );

            var ok = new Button
            {
                Left = 84,
                Top = 112,
                Width = 100,
                Height = 30,
                Text = "OK",
                DialogResult = DialogResult.OK
            };
            var cancel = new Button
            {
                Left = 196,
                Top = 112,
                Width = 100,
                Height = 30,
                Text = "Отмена",
                DialogResult = DialogResult.Cancel
            };
            Controls.Add( ok );
            Controls.Add( cancel );

            AcceptButton = ok;
            CancelButton = cancel;
        }
    }
}
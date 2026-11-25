using System;
using System.Drawing;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Forms;

namespace TelegramShareAddin
{
    /// <summary>
    /// Небольшое окно «Отправляю…» на время работы с мессенджером.
    ///
    /// Пока оно открыто, поток Office свободен: сообщения обрабатываются, и Word не
    /// переходит в состояние «Не отвечает». Сама отправка идёт в отдельном STA-потоке —
    /// этого требуют буфер обмена и SendKeys.
    /// </summary>
    internal sealed class SendingForm : Form
    {
        private const int WmSysCommand = 0x0112;
        private const int ScClose = 0xF060;

        private bool _finished;

        /// <summary>Работа завершилась до показа окна — закрыться сразу после OnShown.</summary>
        public bool FinishRequested { get; set; }

        public SendingForm( string text )
        {
            Text = "Telegram Share";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size( 420 , 96 );
            ShowInTaskbar = false;
            TopMost = true;
            MinimizeBox = false;
            MaximizeBox = false;

            var label = new Label
            {
                Text = text + Environment.NewLine + Environment.NewLine +
                       "Не переключайте окна — клавиши уйдут в активное окно." ,
                AutoSize = false ,
                Dock = DockStyle.Top ,
                Height = 58 ,
                TextAlign = ContentAlignment.MiddleLeft ,
                Padding = new Padding( 12 , 8 , 12 , 0 )
            };

            var bar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee ,
                MarqueeAnimationSpeed = 30 ,
                Dock = DockStyle.Bottom ,
                Height = 14
            };

            Controls.Add( label );
            Controls.Add( bar );
        }

        /// <summary>Закрывает окно по завершении работы (вызывать в UI-потоке).</summary>
        public void MarkFinished( )
        {
            _finished = true;

            if ( Visible )
            {
                Close( );
            }
            else
            {
                FinishRequested = true;
            }
        }

        protected override void OnShown( EventArgs e )
        {
            base.OnShown( e );

            // Работа могла закончиться до первого показа окна.
            if ( FinishRequested || _finished )
            {
                BeginInvoke( (Action) Close );
            }
        }

        protected override void OnFormClosing( FormClosingEventArgs e )
        {
            // Пока идёт отправка, окно не закрываем: иначе работа продолжится «вслепую».
            if ( !_finished )
            {
                e.Cancel = true;
                return;
            }

            base.OnFormClosing( e );
        }

        protected override void WndProc( ref Message m )
        {
            if ( m.Msg == WmSysCommand && ( m.WParam.ToInt32( ) & 0xFFF0 ) == ScClose )
            {
                return; // кнопку закрытия игнорируем
            }

            base.WndProc( ref m );
        }
    }

    /// <summary>Запуск операции с UI Automation в отдельном STA-потоке с окном прогресса.</summary>
    internal static class SendRunner
    {
        private static int _busy;

        public static void Run( string caption , Action work )
        {
            if ( work == null )
            {
                return;
            }

            // Две отправки одновременно боролись бы за буфер обмена и активное окно.
            if ( Interlocked.CompareExchange( ref _busy , 1 , 0 ) != 0 )
            {
                MessageBox.Show(
                    "Отправка уже выполняется — дождитесь её завершения." ,
                    "Telegram Share" ,
                    MessageBoxButtons.OK ,
                    MessageBoxIcon.Information );
                return;
            }

            Exception failure = null;

            try
            {
                using ( var form = new SendingForm( caption ) )
                {
                    var thread = new Thread( ( ) =>
                    {
                        try
                        {
                            work( );
                        }
                        catch ( Exception ex )
                        {
                            failure = ex;
                        }
                        finally
                        {
                            try
                            {
                                if ( form.IsHandleCreated && !form.IsDisposed )
                                {
                                    form.BeginInvoke( (Action) form.MarkFinished );
                                }
                                else
                                {
                                    form.FinishRequested = true;
                                }
                            }
                            catch
                            {
                            }
                        }
                    } );

                    thread.IsBackground = true;
                    thread.SetApartmentState( ApartmentState.STA );

                    // Поток стартует только после создания окна: иначе BeginInvoke из
                    // рабочего потока бросит исключение, и окно останется висеть.
                    form.Shown += ( sender , args ) => thread.Start( );
                    form.ShowDialog( );

                    if ( !thread.Join( 2000 ) )
                    {
                        AddinLog.Warn( "SendRunner: рабочий поток не завершился за 2 с" );
                    }
                }

                if ( failure != null )
                {
                    ExceptionDispatchInfo.Capture( failure ).Throw( );
                }
            }
            finally
            {
                Interlocked.Exchange( ref _busy , 0 );
            }
        }
    }
}

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using stdole;

namespace TelegramShareAddin
{
    /// <summary>Преобразование Image в IPictureDisp для картинок ленты Office.</summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter( )
            : base( string.Empty )
        {
        }

        public static IPictureDisp ToPictureDisp( Image image )
        {
            return (IPictureDisp)GetIPictureDispFromPicture( image );
        }
    }

    /// <summary>Рисование значков кнопок ленты.</summary>
    internal static class IconFactory
    {
        private static readonly Color Blue = Color.FromArgb( 42 , 171 , 238 );

        /// <summary>Прямоугольный значок.</summary>
        public static Bitmap CreateRectangular( )
        {
            return Draw( false );
        }

        /// <summary>Скруглённый значок.</summary>
        public static Bitmap CreateRounded( )
        {
            return Draw( true );
        }

        private static Bitmap Draw( bool rounded )
        {
            var bmp = new Bitmap( 32 , 32 );
            using ( var g = Graphics.FromImage( bmp ) )
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear( Color.Transparent );

                using ( var brush = new SolidBrush( Blue ) )
                {
                    if ( rounded )
                    {
                        using ( var path = RoundedRect( new Rectangle( 1 , 1 , 30 , 30 ) , 8 ) )
                        {
                            g.FillPath( brush , path );
                        }
                    }
                    else
                    {
                        g.FillRectangle( brush , new Rectangle( 1 , 1 , 30 , 30 ) );
                    }
                }

                using ( var white = new SolidBrush( Color.White ) )
                {
                    var plane = new Point[ ]
                    {
                        new Point( 7 , 16 ),
                        new Point( 26 , 8 ),
                        new Point( 15 , 25 ),
                        new Point( 19 , 15 )
                    };
                    g.FillPolygon( white , plane );
                }
            }
            return bmp;
        }

        private static GraphicsPath RoundedRect( Rectangle r , int radius )
        {
            var path = new GraphicsPath( );
            int d = radius * 2;
            path.AddArc( r.X , r.Y , d , d , 180 , 90 );
            path.AddArc( r.Right - d , r.Y , d , d , 270 , 90 );
            path.AddArc( r.Right - d , r.Bottom - d , d , d , 0 , 90 );
            path.AddArc( r.X , r.Bottom - d , d , d , 90 , 90 );
            path.CloseFigure( );
            return path;
        }
    }
}
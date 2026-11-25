using System;
using System.IO;
using System.Text;
using System.Threading;

namespace TelegramShareAddin
{
    /// <summary>
    /// Диагностический лог надстройки.
    ///
    /// Пишется в %LOCALAPPDATA%\TelegramShareAddin\addin.log (отдельно от папки установки,
    /// чтобы удаление надстройки не уносило данные). Запись защищена именованным мьютексом:
    /// Word, Excel и Outlook пишут в один файл из разных процессов. Файл ротируется,
    /// чтобы не расти безгранично.
    /// </summary>
    internal static class AddinLog
    {
        private const string MutexName = "Local\\TelegramShareAddinLog";
        private const long MaxBytes = 1024 * 1024;
        private const int MaxFiles = 3;

        public static string DirectoryPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath( Environment.SpecialFolder.LocalApplicationData ) ,
                    Settings.AppName );
            }
        }

        public static string FilePath
        {
            get
            {
                return Path.Combine( DirectoryPath , "addin.log" );
            }
        }

        public static void Info( string message )
        {
            Write( "INFO " , message );
        }

        public static void Warn( string message )
        {
            Write( "WARN " , message );
        }

        public static void Error( string message )
        {
            Write( "ERROR" , message );
        }

        public static void Error( string message , Exception exception )
        {
            Write( "ERROR" , message + " — " + Describe( exception ) );
        }

        /// <summary>Подробности исключения без стека (стек нужен редко, а файл он раздувает).</summary>
        public static string Describe( Exception exception )
        {
            if ( exception == null )
            {
                return "неизвестная ошибка";
            }

            return exception.GetType( ).Name + ": " + exception.Message;
        }

        private static void Write( string level , string message )
        {
            Mutex mutex = null;

            try
            {
                mutex = new Mutex( false , MutexName );
                if ( !mutex.WaitOne( 500 ) )
                {
                    return;
                }

                Directory.CreateDirectory( DirectoryPath );
                Rotate( );

                string line = DateTime.Now.ToString( "yyyy-MM-dd HH:mm:ss.fff" ) + "  " + level + "  " +
                              message + Environment.NewLine;

                using ( var stream = new FileStream( FilePath , FileMode.Append , FileAccess.Write , FileShare.ReadWrite ) )
                using ( var writer = new StreamWriter( stream , new UTF8Encoding( false ) ) )
                {
                    writer.Write( line );
                }
            }
            catch
            {
                // Лог не должен ломать работу надстройки.
            }
            finally
            {
                if ( mutex != null )
                {
                    try
                    {
                        mutex.ReleaseMutex( );
                    }
                    catch
                    {
                    }

                    mutex.Dispose( );
                }
            }
        }

        private static void Rotate( )
        {
            try
            {
                var file = new FileInfo( FilePath );
                if ( !file.Exists || file.Length < MaxBytes )
                {
                    return;
                }

                string oldest = Path.Combine( DirectoryPath , "addin." + MaxFiles + ".log" );
                if ( File.Exists( oldest ) )
                {
                    File.Delete( oldest );
                }

                for ( int i = MaxFiles - 1 ; i >= 1 ; i-- )
                {
                    string from = Path.Combine( DirectoryPath , "addin." + i + ".log" );
                    string to = Path.Combine( DirectoryPath , "addin." + ( i + 1 ) + ".log" );
                    if ( File.Exists( from ) )
                    {
                        if ( File.Exists( to ) )
                        {
                            File.Delete( to );
                        }
                        File.Move( from , to );
                    }
                }

                if ( File.Exists( FilePath ) )
                {
                    File.Move( FilePath , Path.Combine( DirectoryPath , "addin.1.log" ) );
                }
            }
            catch
            {
            }
        }
    }
}

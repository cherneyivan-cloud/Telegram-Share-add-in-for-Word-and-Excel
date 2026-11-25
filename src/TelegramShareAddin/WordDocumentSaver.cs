using System;
using System.IO;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace TelegramShareAddin
{
    /// <summary>
    /// Сохранение документа Word во временный файл.
    ///
    /// Документ хоста не пересохраняется: используются SaveCopyAs, копирование файла
    /// или экспорт в PDF. SaveAs2 здесь неприменим — он меняет путь живого документа
    /// и сохраняет несохранённые правки пользователя на диск.
    /// </summary>
    internal static class WordDocumentSaver
    {
        public static string Save( Word.Document doc , string path , SendFileFormat format )
        {
            if ( format == SendFileFormat.Pdf )
            {
                return SavePdf( doc , path );
            }

            return SaveNativeCopy( doc , path );
        }

        private static string SavePdf( Word.Document doc , string path )
        {
            try
            {
                doc.ExportAsFixedFormat( path , Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF );
            }
            catch ( Exception ex )
            {
                throw new InvalidOperationException( "Не удалось экспортировать документ в PDF: " + ex.Message , ex );
            }

            EnsureFile( path );
            return path;
        }

        private static string SaveNativeCopy( Word.Document doc , string path )
        {
            // 1. Штатный путь: копия без изменения документа. Повторяем: файл может
            //    быть кратковременно занят антивирусом или индексатором.
            for ( int attempt = 0 ; attempt < 3 ; attempt++ )
            {
                try
                {
                    object fileName = path;
                    doc.SaveCopyAs( ref fileName );

                    if ( IsUsable( path ) )
                    {
                        return path;
                    }
                }
                catch ( Exception ex )
                {
                    AddinLog.Warn( "Word SaveCopyAs (попытка " + ( attempt + 1 ) + "): " + AddinLog.Describe( ex ) );
                }

                Thread.Sleep( 300 );
            }

            // 2. Документ сохранён и не имеет несохранённых правок — файл можно
            //    скопировать как есть, содержимое будет точным.
            try
            {
                string original = doc.FullName;
                bool saved = doc.Saved;

                if ( saved && !string.IsNullOrEmpty( original ) && File.Exists( original ) )
                {
                    File.Copy( original , path , true );

                    if ( IsUsable( path ) )
                    {
                        AddinLog.Info( "Word: отправлена файловая копия документа" );
                        return path;
                    }
                }
            }
            catch ( Exception ex )
            {
                AddinLog.Warn( "Word File.Copy: " + AddinLog.Describe( ex ) );
            }

            // 3. Резервный путь: PDF всегда отражает текущее состояние и не трогает
            //    оригинал. Формат меняется, поэтому сообщаем об этом в лог.
            string pdfPath = Path.ChangeExtension( path , ".pdf" );

            try
            {
                doc.ExportAsFixedFormat( pdfPath , Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF );

                if ( IsUsable( pdfPath ) )
                {
                    AddinLog.Warn( "Word: DOCX-копия недоступна, документ отправлен в PDF (" + pdfPath + ")" );
                    return pdfPath;
                }
            }
            catch ( Exception ex )
            {
                AddinLog.Error( "Word PDF fallback" , ex );
            }

            throw new InvalidOperationException(
                "Не удалось подготовить копию документа. Сохраните его (Ctrl+S) и повторите либо выберите формат PDF." );
        }

        private static bool IsUsable( string path )
        {
            try
            {
                return File.Exists( path ) && ( new FileInfo( path ) ).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static void EnsureFile( string path )
        {
            if ( !IsUsable( path ) )
            {
                throw new InvalidOperationException( "Временный файл не создан или пуст: " + path );
            }
        }
    }
}

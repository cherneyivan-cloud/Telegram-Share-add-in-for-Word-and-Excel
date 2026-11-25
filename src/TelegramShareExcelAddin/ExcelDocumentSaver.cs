using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TelegramShareAddin
{
    /// <summary>Сохранение книги Excel во временный файл (PDF или копия в родном формате).</summary>
    internal static class ExcelDocumentSaver
    {
        public static void Save( Excel.Workbook workbook , string path , SendFileFormat format )
        {
            if ( format == SendFileFormat.Pdf )
            {
                try
                {
                    workbook.ExportAsFixedFormat( Excel.XlFixedFormatType.xlTypePDF , path );
                }
                catch ( Exception ex )
                {
                    throw new InvalidOperationException( "Не удалось экспортировать книгу в PDF: " + ex.Message , ex );
                }
            }
            else
            {
                try
                {
                    workbook.SaveCopyAs( path );
                }
                catch ( Exception ex )
                {
                    AddinLog.Warn( "Excel SaveCopyAs: " + AddinLog.Describe( ex ) );

                    // Копирование файла с диска отправило бы последнюю сохранённую версию
                    // без несохранённых изменений — молча подменять содержимое нельзя.
                    throw new InvalidOperationException(
                        "Не удалось сохранить копию книги. Сохраните её (Ctrl+S) и повторите либо выберите формат PDF." , ex );
                }
            }

            try
            {
                if ( !File.Exists( path ) || ( new FileInfo( path ) ).Length == 0 )
                {
                    throw new InvalidOperationException( "Временный файл не создан или пуст: " + path );
                }
            }
            catch ( InvalidOperationException )
            {
                throw;
            }
            catch
            {
            }
        }
    }
}

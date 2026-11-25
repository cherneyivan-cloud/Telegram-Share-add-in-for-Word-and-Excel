using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace TelegramShareAddin
{
    /// <summary>
    /// Сохранение документа Word во временный файл (PDF или копия в родном формате).
    /// Возвращает фактический путь к подготовленному файлу.
    ///
    /// Документ хоста не пересохраняется: используется только SaveCopyAs/экспорт.
    /// SaveAs2 здесь неприменим — он меняет путь живого документа и сохраняет
    /// несохранённые правки пользователя на диск.
    /// </summary>
    internal static class WordDocumentSaver
    {
        public static string Save( Word.Document doc , string path , SendFileFormat format )
        {
            if ( format == SendFileFormat.Pdf )
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

            return SaveDocxCopy( doc , path );
        }

        private static string SaveDocxCopy( Word.Document doc , string path )
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
                AddinLog.Warn( "Word SaveCopyAs: " + AddinLog.Describe( ex ) );
            }

            throw new InvalidOperationException(
                "Не удалось сделать копию документа. Сохраните его (Ctrl+S) и повторите либо выберите формат PDF." );
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

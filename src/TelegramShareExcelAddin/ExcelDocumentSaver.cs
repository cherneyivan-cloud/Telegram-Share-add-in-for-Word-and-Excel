using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TelegramShareAddin
{
    /// <summary>Сохранение книги Excel во временный файл (PDF или XLSX(.</summary>
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
                    AddinCore.Log( "Excel SaveCopyAs ERROR: " + ex.ToString( ) );

                    try
                    {
                        string full = workbook.FullName;
                        if ( !string.IsNullOrEmpty( full ) && File.Exists( full ) )
                        {
                            File.Copy( full , path , true );
                        }
                    }
                    catch ( Exception ex2 )
                    {
                        AddinCore.Log( "Excel File.Copy ERROR: " + ex2.ToString( ) );
                    }
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
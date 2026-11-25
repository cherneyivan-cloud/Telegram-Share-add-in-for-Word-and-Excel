using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace TelegramShareAddin
{
    /// <summary>
    /// Сохранение документа Word во временный файл (PDF или DOCX).
    /// Возвращает фактический путь к подготовленному файлу.
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
            // 1) Прямое копирование (для документов формата .docx/.docm).
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
                AddinCore.Log( "SaveCopyAs ERROR: " + ex.Message );
            }

            string original = null;
            try
            {
                original = doc.FullName;
            }
            catch
            {
            }

            bool hadPath = !string.IsNullOrEmpty( original ) && original.Length > 2 && original[1] == ':';

            // 2) Конвертация в .docx через SaveAs2 с восстановлением исходного имени.
            if ( hadPath )
            {
                try
                {
                    object outPath = path;
                    object outFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
                    doc.SaveAs2( ref outPath , ref outFormat );

                    bool ok = IsUsable( path );

                    try
                    {
                        object restore = original;
                        doc.SaveAs2( ref restore );
                    }
                    catch ( Exception ex )
                    {
                        AddinCore.Log( "Восстановление имени документа: " + ex.Message );
                    }

                    if ( ok )
                    {
                        return path;
                    }
                }
                catch ( Exception ex )
                {
                    AddinCore.Log( "SaveAs2 (docx) ERROR: " + ex.Message );
                }
            }

            // 3) Копия сохранённого файла с его настоящим расширением.
            if ( hadPath )
            {
                try
                {
                    string ext = Path.GetExtension( original );
                    if ( string.IsNullOrEmpty( ext ) )
                    {
                        ext = ".docx";
                    }

                    string alt = Path.ChangeExtension( path , ext );
                    File.Copy( original , alt , true );
                    if ( IsUsable( alt ) )
                    {
                        AddinCore.Log( "Сохранена файловая копия: " + alt );
                        return alt;
                    }
                }
                catch ( Exception ex )
                {
                    AddinCore.Log( "File.Copy ERROR: " + ex.Message );
                }
            }

            throw new InvalidOperationException( "Не удалось сохранить документ. Сохраните его на диск (Файл -> Сохранить) и повторите, либо выберите формат PDF." );
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
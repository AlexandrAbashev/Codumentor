using System.Collections.ObjectModel;
using System.IO;
using ColorCode;
using Microsoft.Office.Interop.Word;
using Languages = ColorCode.Languages;
using Task = System.Threading.Tasks.Task;
using StyleSheets = ColorCode.StyleSheets;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;
using System.Runtime.InteropServices;
using System.Text;

namespace Codumentor.Services
{
    public class CodeToWordExporter
    {
        public void ExportCodeToDocument(ObservableCollection<string> filePaths, string outputFilePath, bool isPDF = true)
        {
            var _wordApp = new Application();
            var doc = _wordApp.Documents.Add();
            try
            {
                foreach (var filePath in filePaths)
                {
                    if (File.Exists(filePath))
                    {
                        string fileName = Path.GetFileName(filePath);
                        string code = File.ReadAllText(filePath);

                        InsertFileName(doc, fileName);

                        string fileExtension = Path.GetExtension(filePath).ToLower();

                        InsertHighlightedCode(doc, code, fileExtension);
                    }
                }
                if (isPDF)
                    doc.SaveAs2(outputFilePath, WdSaveFormat.wdFormatPDF);
                else
                    doc.SaveAs2(outputFilePath);
            }
            catch (UnauthorizedAccessException)
            {
                throw new InvalidOperationException("Нет прав на запись файла. Попробуйте выбрать другую папку.");
            }
            catch (IOException)
            {
                throw new InvalidOperationException("Файл занят другим процессом или нет доступа к папке.");
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"COM ошибка: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }
            finally
            {
                doc?.Close(false);
                Marshal.ReleaseComObject(doc);
                _wordApp.Quit();
            }
        }

        public Task ExportCodeToDocumentAsync(ObservableCollection<string> filePaths, string outputPath, bool isPDF = true)
        {
            return Task.Run(() => ExportCodeToDocument(filePaths, outputPath));
        }

        private void InsertFileName(Document doc, string fileName)
        {
            Range fileNameInsertRange = doc.Content;
            fileNameInsertRange.Collapse(WdCollapseDirection.wdCollapseEnd);

            int fileNameStart = fileNameInsertRange.Start;
            fileNameInsertRange.InsertAfter(fileName + "\r\n");
            Range fileNameRange = doc.Range(fileNameStart, fileNameInsertRange.End);
            fileNameRange.Font.Bold = 1;
            fileNameRange.Font.Size = 12;
            fileNameRange.Font.Name = "Times New Roman";
            fileNameRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private void InsertHighlightedCode(Document doc, string code, string fileExtension)
        {
            string highlightedCode = HighLightCode(code, GetLanguageByExtension(fileExtension));

            string fullHtml = $"<html><body>{highlightedCode}</body></html>";

            string tempHtmlPath = Path.Combine(Path.GetTempPath(), "temp_highlighted.html");
            File.WriteAllText(tempHtmlPath, fullHtml, Encoding.UTF8);

            try
            {
                Range insertRange = doc.Content;
                insertRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertFile(tempHtmlPath);
            }
            finally
            {
                if (File.Exists(tempHtmlPath))
                    File.Delete(tempHtmlPath);
            }
        }

        private string HighLightCode(string code, ILanguage language)
        {
            var colorizer = new CodeColorizer();

            if (language != null)
            {
                using (var writer = new StringWriter())
                {
                    colorizer.Colorize(
                        sourceCode: code,
                        language: language,
                        textWriter: writer,
                        formatter: Formatters.Default,
                        styleSheet: StyleSheets.Default
                    );
                    return writer.ToString();
                }
            }
            return code;
        }

        private ILanguage? GetLanguageByExtension(string fileExtension)
        {
            ILanguage? language = null;
            switch (fileExtension)
            {
                case ".cs":
                    language = Languages.CSharp;
                    break;
                case ".java":
                    language = Languages.Java;
                    break;
                case ".php":
                    language = Languages.Php;
                    break;
                case ".js":
                    language = Languages.JavaScript;
                    break;
                case ".cpp":
                    language = Languages.Cpp;
                    break;
                case ".css":
                    language = Languages.Css;
                    break;
                case ".xaml":
                case ".xml":
                    language = Languages.Xml;
                    break;
            }
            return language;
        }
    }
}

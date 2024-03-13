using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.IO;

namespace COM
{
    internal class Program
    {
        static void Main()
        {
            // Открываем приложение Word
            Word.Application wordApp = new Word.Application();
            // wordApp.Visible = true; // Отобразить окно приложения

            // Относительные пути высчитываются от папки Документы, а не от папки с раположением .exe файла
            var localPath = @".\temp.docx";
            object fullPath = Path.GetFullPath(localPath);

            Console.WriteLine();
            // Открываем документ
            Word.Document wDoc = wordApp.Documents.Add(ref fullPath, false, Word.WdNewDocumentType.wdNewBlankDocument, true);
            // Заменяем слова-теги на нужные значения
            Replace("{СОБЫТИЕ}", "извержение вулкана");
            Replace("{ДЕЙСТВИЕ}", "привлекло внимание");
            Replace("{СОБЫТИЕ}", "извержение вулкана");
            Replace("{РЕАКЦИЯ}", "панику");
            Replace("{НАБЛЮДАТЕЛИ}", "местные жители");
            Replace("{КОЛ-ВО РАЗ}", "два раза");
            Replace("{РЕЗУЛЬТАТ ДЕЙСТВИЯ}", "необходимость эвакуации");
            // Редактируем содержимое закладки
            wDoc.Bookmarks["mark"].Range.Text = "Тут была закладка";

            try
            {
                // Сохраняем файл (если не указать расширение, то будет использоваться .docx)
                wDoc.SaveAs2(@".\output.docx");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            // Закрываем приложение 
            wordApp.Quit(Word.WdSaveOptions.wdPromptToSaveChanges);

            void Replace(string find, string replace)
            {
                // Выбираем весь текст
                Word.Range range = wDoc.StoryRanges[Word.WdStoryType.wdMainTextStory];
                range.Find.ClearFormatting();
                // Используем встроенную функцию замены
                range.Find.Execute(FindText: find, ReplaceWith: replace);
            }

            // Запускаем приложение и создаём новую книгу таблиц
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook book = excelApp.Workbooks.Add();
            Excel.Worksheet ws = book.ActiveSheet;

            for (int i = 1, num = 60; i < 10; i++, num--)
            {
                ws.Cells[1, i] = i;
                ws.Cells[2, i] = num;
            }
            Excel.Range cell = ws.Cells[3, 1];
            cell.Formula = "=SUM(A1:J1)";
            cell.FormulaHidden = false;
            cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Создаём график и сохраняем в переменную. Если объект графика не будет сохранён в переменную, то он будет высвобожден из памяти и удалён
            Excel.ChartObject myChart = ((Excel.ChartObjects)ws.ChartObjects(Type.Missing)).Add(50, 100, 400, 300);
            Excel.Chart chart = myChart.Chart;
            chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;
            Excel.Series series = ((Excel.SeriesCollection)chart.SeriesCollection(Type.Missing)).NewSeries();
            series.XValues = ws.Range["A1:J1"];
            // Испоьлзуется вместо chart.Values
            chart.SetSourceData(ws.Range["A2:J1"]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = "График из C#";
            chart.HasLegend = true;
            series.Name = "Title";

            try
            {
                ws.SaveAs2(@".\output.xls");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            excelApp.Quit();
        }
    }
}

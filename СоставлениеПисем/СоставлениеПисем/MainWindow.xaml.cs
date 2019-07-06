using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace СоставлениеПисем
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenFile(TextBlock TextBlockFileName,string Filter,string DefaulFilter)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = DefaulFilter; // Default file extension
            dlg.Filter = Filter; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                TextBlockFileName.Text = dlg.FileName;
            }
        }
        private void SaveAsFile(Word.Document wordFile)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.DefaultExt = ".doc"; // Default file extension
            dlg.Filter = "Word files (.doc)|*.doc"; // Filter files by extension
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
            {
                // получаем выбранный файл
                string filename = dlg.FileName;

                if (System.IO.Path.GetExtension(dlg.FileName).ToLower() == ".doc")
                    wordFile.SaveAs2(dlg.FileName, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument97);
                else // Assuming a "docx" extension:
                    wordFile.SaveAs2(dlg.FileName);
            }
        }
        private void ButtonOpenDataFile_Click(object sender, RoutedEventArgs e)
        {
            // открыть файл Excel с данными
            OpenFile(TextBlockDataFilePatch, "Excel files (.xls)|*.xls", ".xls");
        }
        private void ButtonOpenTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            // открыть файл-шаблон пригласительного письма
            OpenFile(TextBlockTemplateFilePatch, "Word files (.doc)|*.doc", ".doc");
        }
        private void ButtonOpenTemplateLetterFile_Click(object sender, RoutedEventArgs e)
        {
            // открыть файл-шаблон для печати на конверте
            OpenFile(TextBlockTemplateLetterPatch, "Word files (.doc)|*.doc", ".doc");
        }
        private void ButtonPrintLetter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (TextBlockDataFilePatch.Text == "")
                    return;
                if (TextBlockTemplateFilePatch.Text == "")
                    return;
                FileInfo fiDataFilePatch = new FileInfo(TextBlockDataFilePatch.Text);
                FileInfo fiTemplateFilePatch = new FileInfo(TextBlockTemplateFilePatch.Text);
                // Проверяем, что выбранные файлы существуют
                if (fiDataFilePatch.Exists && fiTemplateFilePatch.Exists)
                {
                    // Открываем файл Excel с данными
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wbv = excel.Workbooks.Open(TextBlockDataFilePatch.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet wx = (Excel.Worksheet)wbv.Worksheets.get_Item(1);
                    // word - файл с шаблоном пригласительного письма, 
                    // в нем далее будут формироваться (добавляться в конец) 
                    // письма по всем компаниям
                    Word._Application word = new Word.Application();
                    // word2 - файл с шаблоном пригласительного письма, 
                    // который используется единожды для заполнения
                    // данных по конкретной компании, после чего
                    // листы добавляются в конец word
                    Word._Application word2 = new Word.Application();
                    // Открываем файл  шаблоном пригласительного письма
                    Word.Document wordFile = word.Documents.Open(TextBlockTemplateFilePatch.Text);
                    // Номер строки, в которой указана первая компания
                    int i = Convert.ToInt32(TextBoxNumberFrom.Text.ToString());
                    // Заполняются данные по первой компании
                    wordFile.FormFields["ТекстовоеПоле1"].Range.Text = wx.Cells[i, 4].Text;
                    wordFile.FormFields["ТекстовоеПоле2"].Range.Text = wx.Cells[i, 2].Text;
                    wordFile.FormFields["ТекстовоеПоле3"].Range.Text = wx.Cells[i, 8].Text;
                    wordFile.FormFields["ТекстовоеПоле4"].Range.Text = wx.Cells[i, 3].Text;
                    wordFile.FormFields["ТекстовоеПоле5"].Range.Text = wx.Cells[i, 6].Text;
                    wordFile.FormFields["ТекстовоеПоле6"].Range.Text = wx.Cells[i, 7].Text;
                    // В конец файла добавляется новый пустой абзац
                    object oEndofDoc = "\\endofdoc";
                    object oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                    Word.Paragraph par = wordFile.Content.Paragraphs.Add(ref oRng);
                    // В конец файла добавляется разрыв страницы
                    object unit;
                    object extend;
                    unit = Word.WdUnits.wdStory;
                    extend = Word.WdMovementType.wdMove;
                    word.Selection.EndKey(ref unit, ref extend);
                    object oType;
                    oType = Word.WdBreakType.wdSectionBreakNextPage;
                    word.Selection.InsertBreak(ref oType);
                    // После разрыва страницы снова добавляется пустой абзац
                    oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                    par = wordFile.Content.Paragraphs.Add(ref oRng);

                    object missing = System.Reflection.Missing.Value;
                    object readOnly = true;
                    // Запоминаем номер строки, до которой необходимо вывести данные
                    int N = Convert.ToInt32(TextBoxNumberBefor.Text.ToString());
                    for (i++; i <= N; i++)
                    {
                        // Открываем файл с шаблоном пригласительного письма
                        Word._Document oDoc = word2.Documents.Open(TextBlockTemplateFilePatch.Text, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        Word.Range oRange = oDoc.Content;
                        // Заполняются данные по первой компании
                        oDoc.FormFields["ТекстовоеПоле1"].Range.Text = wx.Cells[i, 4].Text;
                        oDoc.FormFields["ТекстовоеПоле2"].Range.Text = wx.Cells[i, 2].Text;
                        oDoc.FormFields["ТекстовоеПоле3"].Range.Text = wx.Cells[i, 8].Text;
                        oDoc.FormFields["ТекстовоеПоле4"].Range.Text = wx.Cells[i, 3].Text;
                        oDoc.FormFields["ТекстовоеПоле5"].Range.Text = wx.Cells[i, 6].Text;
                        oDoc.FormFields["ТекстовоеПоле6"].Range.Text = wx.Cells[i, 7].Text;
                        // Копируем шаблон с заполненными данными в буфер обмена
                        oRange.Copy();
                        // Вставляем из буфера обмена страницы в конец wordFile
                        par.Range.Paste();
                        // В конец файла wordFile добавляется разрыв страницы
                        unit = Word.WdUnits.wdStory;
                        extend = Word.WdMovementType.wdMove;
                        word.Selection.EndKey(ref unit, ref extend);
                        oType = Word.WdBreakType.wdSectionBreakNextPage;
                        word.Selection.InsertBreak(ref oType);
                        // В конец файла wordFile добавляется разрыв страницы
                        oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                        par = wordFile.Content.Paragraphs.Add(ref oRng);
                        // Закрываем файл oDoc 
                        oDoc.Close(false, Type.Missing, Type.Missing);
                    }
                    // Показывваем файл со всеми заполненными пригласительными письмами
                    word.Visible = true;
                    // Сохраняем файл
                    //SaveAsFile(wordFile);
                    //wordFile.Close();
                    // Закрываем исходный файл Excel с данными
                    wbv.Close(false, Type.Missing, Type.Missing);
                    excel.Quit();
                }
                else
                {
                    MessageBox.Show("Выберитие файл шаблона и файл с данными");
                }
            }
            catch (Exception ex)
            {
                // Показываем сообщение об ошибке
                MessageBox.Show(ex.Message);
            }
        }
        private void ButtonPrintLetterBox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (TextBlockDataFilePatch.Text == "")
                    return;
                if (TextBlockTemplateLetterPatch.Text == "")
                    return;
                FileInfo fiDataFilePatch = new FileInfo(TextBlockDataFilePatch.Text);
                FileInfo fiTemplateFilePatch = new FileInfo(TextBlockTemplateLetterPatch.Text);
                if (fiDataFilePatch.Exists && fiTemplateFilePatch.Exists)
                {

                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook wbv = excel.Workbooks.Open(TextBlockDataFilePatch.Text, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet wx = (Excel.Worksheet)wbv.Worksheets.get_Item(1);

                    Word._Application word = new Word.Application();
                    Word._Application word2 = new Word.Application();
                    //oDoc.Activate();

                    Word.Document wordFile = word.Documents.Open(TextBlockTemplateLetterPatch.Text); // word.Documents.Add();

                    int i = Convert.ToInt32(TextBoxNumberFrom.Text.ToString());
                    wordFile.FormFields["ТекстовоеПоле1"].Range.Text = wx.Cells[i, 4].Text;
                    wordFile.FormFields["ТекстовоеПоле2"].Range.Text = wx.Cells[i, 2].Text;
                    wordFile.FormFields["ТекстовоеПоле3"].Range.Text = wx.Cells[i, 8].Text;
                    wordFile.FormFields["ТекстовоеПоле4"].Range.Text = wx.Cells[i, 13].Text;
                    wordFile.FormFields["ТекстовоеПоле5"].Range.Text = wx.Cells[i, 12].Text;
                    wordFile.FormFields["ТекстовоеПоле6"].Range.Text = wx.Cells[i, 13].Text;

                    object oEndofDoc = "\\endofdoc";
                    object oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                    Word.Paragraph par = wordFile.Content.Paragraphs.Add(ref oRng);

                    object unit;
                    object extend;
                    unit = Word.WdUnits.wdStory;
                    extend = Word.WdMovementType.wdMove;
                    word.Selection.EndKey(ref unit, ref extend);
                    object oType;
                    oType = Word.WdBreakType.wdSectionBreakNextPage;
                    word.Selection.InsertBreak(ref oType);

                    oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                    par = wordFile.Content.Paragraphs.Add(ref oRng);

                    object missing = System.Reflection.Missing.Value;
                    object readOnly = true;
                    //word.Visible = true;
                    int N = Convert.ToInt32(TextBoxNumberBefor.Text.ToString());
                    for (i++; i <= N; i++)
                    {

                        Word._Document oDoc = word2.Documents.Open(TextBlockTemplateLetterPatch.Text, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        Word.Range oRange = oDoc.Content;


                        oDoc.FormFields["ТекстовоеПоле1"].Range.Text = wx.Cells[i, 4].Text;
                        oDoc.FormFields["ТекстовоеПоле2"].Range.Text = wx.Cells[i, 2].Text;
                        oDoc.FormFields["ТекстовоеПоле3"].Range.Text = wx.Cells[i, 8].Text;
                        oDoc.FormFields["ТекстовоеПоле4"].Range.Text = wx.Cells[i, 13].Text;
                        oDoc.FormFields["ТекстовоеПоле5"].Range.Text = wx.Cells[i, 12].Text;
                        oDoc.FormFields["ТекстовоеПоле6"].Range.Text = wx.Cells[i, 13].Text;

                        oRange.Copy();

                        par.Range.Paste();

                        unit = Word.WdUnits.wdStory;
                        extend = Word.WdMovementType.wdMove;
                        word.Selection.EndKey(ref unit, ref extend);
                        oType = Word.WdBreakType.wdSectionBreakNextPage;
                        word.Selection.InsertBreak(ref oType);

                        oRng = wordFile.Bookmarks.get_Item(ref oEndofDoc).Range;
                        par = wordFile.Content.Paragraphs.Add(ref oRng);


                        oDoc.Close(false, Type.Missing, Type.Missing);
                    }

                    word.Visible = true;

                    //SaveAsFile(wordFile);
                    //wordFile.Close();

                    wbv.Close(false, Type.Missing, Type.Missing);
                    excel.Quit();

                    //word.Quit();
                }
                else
                {
                    MessageBox.Show("Выберитие файл шаблона и файл с данными");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

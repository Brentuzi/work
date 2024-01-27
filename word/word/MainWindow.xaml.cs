using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
namespace word
{
    public partial class MainWindow : Window
    {
        ObservableCollection<Item> items;

        public MainWindow()
        {
            InitializeComponent();
            InitializeData();
            itemGrid.ItemsSource = items;
            itemGrid.RowEditEnding += ItemGrid_RowEditEnding;
            dateLabel.Content = DateTime.Now.ToString("dd.MM.yyyy");
        }

        private void InitializeData()
        {
            items = new ObservableCollection<Item>
            {
                new Item {Name="проверка", Amount=3, Price=16}
             
            };
            items.CollectionChanged += (s, e) => UpdateTotal();
            UpdateTotal();
        }

        private void CreateAndSaveExcelFile()
        {
            var excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Cells[1, 1] = "Название";
            worksheet.Cells[1, 2] = "Количество";
            worksheet.Cells[1, 3] = "Цена";
            worksheet.Cells[1, 4] = "Сумма";

            int row = 2;
            foreach (var item in items)
            {
                worksheet.Cells[row, 1] = item.Name;
                worksheet.Cells[row, 2] = item.Amount;
                worksheet.Cells[row, 3] = item.Price;
                worksheet.Cells[row, 4] = item.Sum;
                row++;
            }


            string savePath = @"C:\Users\429193-25\source\repos\word\exel.xlsx";

            try
            {
                workbook.SaveAs(savePath);
                excelApp.Quit();
                MessageBox.Show($"Excel файл сохранен: {savePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении Excel файла: {ex.Message}");
            }
        }

        private void UpdateTotal()
        {
            double total = items.Sum(x => x.Sum);
            totalLabel.Content = $"Итого: {total} рублей";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CreateWordDocument();
            CreateAndSaveExcelFile();
        }

        private void CreateWordDocument()
        {
            Word.Application app = new Word.Application();
            app.Visible = true;

            Word.Document doc = app.Documents.Add();
            AddInvoiceHeader(doc);
            AddSupplierAndBuyerInfo(doc);
            AddItemsTable(doc);
            AddTotal(doc);
            string savePath = @"C:\Users\429193-25\source\repos\word\word.docx";

            try
            {
                doc.SaveAs2(savePath);
                MessageBox.Show($"Документ сохранен: {savePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении документа: {ex.Message}");
            }
         
        }

        private void AddInvoiceHeader(Word.Document doc)
        {
            Word.Paragraph titlePar = doc.Content.Paragraphs.Add();
            titlePar.Range.Text = $"Накладная №{invIDTextbox.Text} от {DateTime.Now.ToString("dd.MM.yyyy")}";
            titlePar.Range.Font.Name = "Times New Roman";
            titlePar.Range.Font.Size = 14;
            titlePar.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titlePar.Range.InsertParagraphAfter();
        }

        private void AddSupplierAndBuyerInfo(Word.Document doc)
        {
            AddParagraphWithUnderline(doc, $"Поставщик: {supplierTextBox.Text}");
            AddParagraphWithUnderline(doc, $"Покупатель: {buyerTextBox.Text}");
        }

        private void AddItemsTable(Word.Document doc)
        {
            Word.Paragraph tablePar = doc.Content.Paragraphs.Add();
            tablePar.Range.InsertParagraphAfter();

            Word.Table tbl = doc.Tables.Add(tablePar.Range, items.Count + 1, 5);
            tbl.Borders.Enable = 1;

       
            string[] headers = { "№", "Название", "Количество", "Цена", "Сумма" };
            for (int i = 0; i < headers.Length; i++)
            {
                tbl.Rows[1].Cells[i + 1].Range.Text = headers[i];
                tbl.Rows[1].Cells[i + 1].Range.Bold = 1;
                tbl.Rows[1].Cells[i + 1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }

          
            for (int i = 0; i < items.Count; i++)
            {
                tbl.Rows[i + 2].Cells[1].Range.Text = (i + 1).ToString();
                tbl.Rows[i + 2].Cells[2].Range.Text = items[i].Name;
                tbl.Rows[i + 2].Cells[3].Range.Text = items[i].Amount.ToString();
                tbl.Rows[i + 2].Cells[4].Range.Text = items[i].Price.ToString();
                tbl.Rows[i + 2].Cells[5].Range.Text = items[i].Sum.ToString();
            }
        }

        private void AddTotal(Word.Document doc)
        {
            Word.Paragraph ttl = doc.Content.Paragraphs.Add();
            ttl.Range.Text = totalLabel.Content.ToString();
            ttl.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            ttl.Range.Font.Name = "Times New Roman";
            ttl.Range.Font.Size = 14;
            ttl.Range.InsertParagraphAfter();
        }

        private void AddParagraphWithUnderline(Word.Document doc, string text)
        {
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Range.Text = text;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Underline = Word.WdUnderline.wdUnderlineSingle;
            paragraph.Range.Bold = 1;
            paragraph.Range.InsertParagraphAfter();
        }

        private void ItemGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var item = e.Row.Item as Item;
                if (item != null && !items.Contains(item))
                {
                    items.Add(new Item());
                }
            }
            UpdateTotal();
        }
    }

}

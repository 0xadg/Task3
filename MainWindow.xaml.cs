using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ObservableCollection<Item> l;
        Word.Application app;

        public MainWindow ()
        {
            InitializeComponent();
            l = new ObservableCollection<Item>
            {
                new Item {Name="Арбузы", Amount=3, Price=16},
                new Item {Name="Тыква", Amount=10, Price=40},
                new Item {Name="Яблоки", Amount=8, Price=80}
            };
            l.CollectionChanged += addHandlerToNewItem;
            for(int i = 0; i < l.Count; i++)
            {
                l[i].PropertyChanged += updateTotal;

            }
            itemGrid.ItemsSource = l;
            updateTotal(this, null);
            dateLabel.Content = DateTime.Now.ToString("dd.MM.yyyy");
        }

        private void addHandlerToNewItem(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            for(int i = 0; i < e.NewItems.Count; i++)
            {
                ((Item)e.NewItems[i]).PropertyChanged += updateTotal;
            }
        }

        private void updateTotal (object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            double ttl = l.Sum(x => x.Sum);
            totalLabel.Content = $"Итого: {ttl} рублей";
        }
        private void Button_Click (object sender, RoutedEventArgs e)
        {
            app = new Word.Application();
            app.Visible = true;
            

            Word.Document doc = app.Documents.Add();

            Word.Paragraph titlePar = doc.Content.Paragraphs.Add();
            string invoiceID = invIDTextbox.Text;
            titlePar.Range.Text = $"Накладная №{invoiceID} от {DateTime.Now.ToString("dd.MM.yyyy")}";
            titlePar.Range.Font.Name = "Times New Roman";
            titlePar.Range.Font.Size = 14;
            titlePar.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titlePar.Range.InsertParagraphAfter();
            // sets alignment of further paragraphs
            titlePar.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            Word.Paragraph supplierPar = doc.Content.Paragraphs.Add();//addParagraph(doc, $"Поставщик: {supplierTextBox.Text}", "Times New Roman", 12);
            supplierPar.Range.Text = $"Поставщик: {supplierTextBox.Text}";
            supplierPar.Range.Font.Name = "Times New Roman";
            supplierPar.Range.Font.Size = 14;

            object start = supplierPar.Range.Start + 11;
            object end = supplierPar.Range.Start + 11 + buyerTextBox.Text.Length;

            Word.Range rng = doc.Range(start, end);
            rng.Underline = Word.WdUnderline.wdUnderlineSingle;
            rng.Bold = 1;
            rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            supplierPar.Range.InsertParagraphAfter();

            Word.Paragraph buyerPar = doc.Content.Paragraphs.Add();//addParagraph(doc, $"Покупатель: {buyerTextBox.Text}", "Times New Roman", 12);
            buyerPar.Range.Text = $"Покупатель: {buyerTextBox.Text}";
            buyerPar.Range.Font.Name = "Times New Roman";
            buyerPar.Range.Font.Size = 14;

            object start2 = buyerPar.Range.Start + 12;
            object end2 = buyerPar.Range.Start + 12 + buyerTextBox.Text.Length;

            Word.Range rng2 = doc.Range(start2, end2);
            rng2.Underline = Word.WdUnderline.wdUnderlineSingle;
            rng2.Bold = 1;
            rng2.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            buyerPar.Range.InsertParagraphAfter();

            Word.Paragraph tablePar = doc.Content.Paragraphs.Add();
            tablePar.Range.InsertParagraphAfter();
            
            Word.Table tbl = doc.Tables.Add(tablePar.Range, l.Count+1, 5);
            
            tbl.Borders.Enable = 1;

            Word.Cells headers = tbl.Rows[1].Cells;
            headers[1].Range.Text = "№";
            headers[1].Range.Bold = 1;
            headers[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            headers[2].Range.Text = "Название";
            headers[2].Range.Bold = 1;
            headers[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            
            headers[3].Range.Text = "Количество";
            headers[3].Range.Bold = 1;
            headers[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            
            headers[4].Range.Text = "Цена";
            headers[4].Range.Bold = 1;
            headers[4].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            headers[5].Range.Text = "Сумма";
            headers[5].Range.Bold = 1;
            headers[5].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            for (int i = 1; i <= l.Count; i++)
            {
                Word.Cells row = tbl.Rows[i+1].Cells;
                row[1].Range.Text = $"{i}";
                row[2].Range.Text = l[i - 1].Name;
                row[3].Range.Text = l[i - 1].Amount.ToString();
                row[4].Range.Text = l[i - 1].Price.ToString();
                row[5].Range.Text = l[i - 1].Sum.ToString();
            }

            Word.Paragraph ttl = doc.Content.Paragraphs.Add();
            ttl.Range.Text = totalLabel.Content.ToString();
            ttl.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            ttl.Range.Font.Name = "Times New Roman";
            ttl.Range.Font.Size = 14;
            ttl.Range.InsertParagraphAfter();
        }

        private void itemGrid_LoadingRow (object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void TextBox_TextChanged (object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click_1 (object sender, RoutedEventArgs e)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            app.Workbooks.Add();
            app.Range["A2"].Value = "ID";
            app.Range["B2"].Value = "Наименование";
            app.Range["C2"].Value = "Кол-во";
            app.Range["D2"].Value = "Цена";
            app.Range["E2"].Value = "Сумма";

            app.Range["A1:E1"].Merge();
            app.Range["A1"].Value = $"Накладная №{invIDTextbox.Text} от {DateTime.Now.ToString("dd.MM.yyyy")}";
            app.Range["A1"].Font.Bold = 1;
            app.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //app.Range["C1:E1"].Merge();
            //app.Range["C1"].Value = "";

            app.Range["A2"].Font.Bold = 1;
            app.Range["B2"].Font.Bold = 1;
            app.Range["C2"].Font.Bold = 1;
            app.Range["D2"].Font.Bold = 1;
            app.Range["E2"].Font.Bold = 1;

            app.Range["A2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            app.Range["B2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            app.Range["C2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            app.Range["D2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            app.Range["E2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            app.Columns.AutoFit();

            app.Range["A3"].Select();

            for (int i = 1; i <= l.Count; i++)
            {
                app.ActiveCell.Value = $"{i}";
                app.ActiveCell.Offset[0, 1].Value = l[i - 1].Name;
                app.ActiveCell.Offset[0, 2].Value = l[i - 1].Amount.ToString();
                app.ActiveCell.Offset[0, 3].Value = l[i - 1].Price.ToString();
                app.ActiveCell.Offset[0, 4].Value = l[i - 1].Sum.ToString();

                app.ActiveCell.Offset[1, 0].Select();
            }
        }
    }
}

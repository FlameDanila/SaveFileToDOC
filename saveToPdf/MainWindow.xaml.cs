using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace saveToPdf
{ 
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var app = new Word.Application();

            Word.Document document = app.Documents.Add();
            Word.Paragraph userParagraph = document.Paragraphs.Add();
            Word.Range range = userParagraph.Range;

            Word.Table tables = document.Tables.Add(range, 5, 2);

            tables.Borders.InsideLineStyle = tables.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            range = tables.Cell(1, 2).Range;
            range.Text = "text on 1, 2";
            range = tables.Cell(5, 1).Range;
            range.Text = "text on 5,1";

            Word.Paragraph maxParagraf = document.Paragraphs.Add();
            Word.Range maxRange = maxParagraf.Range;
            maxRange.Text = "sdf";
            maxRange.InsertParagraphAfter();

            document.SaveAs2(@"C:\Users\student\Desktop\Build.doc");
            document.SaveAs2(@"C:\Users\student\Desktop\Build.pdf", Word.WdExportFormat.wdExportFormatPDF);
            document.Close();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            //string text = "asdas";
            //using (FileStream file = new FileStream(@"C:\Users\student\Desktop\Build.doc", FileMode.Open))
            //{
            //    using (FileStream save = new FileStream(@"C:\Users\student\Desktop\Build.pdf", FileMode.Create))
            //    {
            //        using (StreamWriter stream = new StreamWriter(save))
            //        {
            //            stream.WriteLine(file);
            //        }
            //    }
            //}   
        }
    }
}

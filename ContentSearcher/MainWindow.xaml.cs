using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
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
using System.Xml;

namespace ContentSearcher
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            //setup
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //trigger

        }

        private void docxSearch(string docLocation)
        {
            
            Uri partUriDocument = PackUriHelper.CreatePartUri(new Uri("word\\document.xml", UriKind.Relative));
            Package package = Package.Open(docLocation, FileMode.Open); //docx megnyitás
            PackagePart packagePartDocument = package.GetPart(partUriDocument);
            Stream xmlStream = packagePartDocument.GetStream(FileMode.Open);
            StringBuilder builder = new StringBuilder();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlStream);
            package.Close();
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"); //doksiból kiemelt schema
            foreach (XmlNode node in xmlDoc.SelectNodes("/descendant::w:p" + "[not(w:r/w:rPr/w:strike)]", nsmgr))
            {
                if (node.InnerXml.Contains("w:b") & node.InnerText.Contains("2.7."))
                {

                }
            }
        }

        private void excelSearch(string xlsLocation)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = app.Workbooks.Open(xlsLocation, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing);
            foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Sheets)
            {

            }
            

        }
    }
}

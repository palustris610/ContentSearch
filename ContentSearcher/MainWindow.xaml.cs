using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ContentSearcher
{
    /* TODO:
     * Függvényt készíteni, ami létrehozza a control-okat
     * -Control-ok neve '_1' számra végződjön, így könnyű beazonosítani melyik sor az
     * -counter-t vezetni, hogy épp hol tartunk
     * -Több szintesítés, zárójelezés kialakítása, fa-szerkezet vagy kapcsos zárójelezés
     * A létező control-okat számba venni, amikor a keresés elindul
     * Több 'szabályos' keresés lebonyolítása - ÉS VAGY a szabályok közt
     * Control sor kitörlése, és igazítása
     * -counter csökkentés, elemek törlése, nameunregister
     * -utána következő elemek ha vannak, akkor átnevezés, mozgatás, nameregister stb
     * Különböző szabályokra kereső függvények
     * 
     */
    public partial class MainWindow : Window
    {
        string textToSearch = string.Empty;
        string defLocation = @"E:\ZZZ_TESZTMAPPA";
        List<string> wordList = new List<string>();
        List<string> excelList = new List<string>();
        List<string> pdfList = new List<string>();
        List<string> outputList = new List<string>();
        BackgroundWorker bw = new BackgroundWorker();

        List<object> logicList = new List<object>();
        List<object> subjectList = new List<object>();
        List<object> operatorList = new List<object>();


        public MainWindow()
        {
            InitializeComponent();

            
            bw.DoWork += Bw_DoWork;
            bw.RunWorkerCompleted += Bw_RunWorkerCompleted;
            ListFillUp();
        }

        private void ListFillUp()
        {
            logicList.Add("AND");
            logicList.Add("OR");
            comboBoxLogic_1.ItemsSource = logicList;
            subjectList.Add("FILE NAME");
            subjectList.Add("FILE CONTENT");
            comboBoxSubject_1.ItemsSource = subjectList;
            operatorList.Add("EQUALS");
            operatorList.Add("NOT EQUALS");
            operatorList.Add("CONTAINS");
            operatorList.Add("NOT CONTAINS");
            operatorList.Add("STARTS WITH");
            operatorList.Add("NOT STARTS WITH");
            operatorList.Add("ENDS WITH");
            operatorList.Add("NOT ENDS WITH");
            comboBoxOperator_1.ItemsSource = operatorList;
        }

        private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Bw_DoWork(object sender, DoWorkEventArgs e)
        {

            //Parallel.ForEach(Directory.GetFiles(e.Argument as string), );
        }

        private void DocSearch()
        {
            try
            {
                Word.Application app = new Word.Application();
                foreach (string file in wordList)
                {
                    Word.Document doc = app.Documents.Open(file, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);
                    foreach (Word.Paragraph parag in doc.Paragraphs)
                    {
                        if (parag.Range.Text.Contains(textToSearch))
                        {
                            listBoxOutput.Items.Add(file);
                        }
                    }
                    doc.Close();
                }
                app.Quit();
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        private void ExcelSearch()
        {
            try
            {
                Excel.Application app = new Excel.Application();
                foreach (string xls in excelList)
                {
                    Excel.Workbook wb = app.Workbooks.Open(xls);
                    foreach (Excel.Worksheet sheet in wb.Sheets)
                    {
                        object missing = Type.Missing;
                        Excel.Range firstFind = null;
                        firstFind = sheet.UsedRange.Find(textToSearch, missing,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
            missing, missing);
                        if (firstFind != null)
                        {
                            
                            listBoxOutput.Items.Add(xls);
                        }
                    }
                    wb.Close();
                }
                app.Quit();
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        private void PdfSearch()
        {
            try
            {
                foreach (string pdf in pdfList)
                {
                    FileStream stream = File.Open(pdf, FileMode.Open);
                    PdfExtract.Extractor extractor = new PdfExtract.Extractor();
                    string temp = extractor.ExtractToString(stream, Encoding.Default);
                    if (temp.Contains(textToSearch))
                    {
                        //outputList.Add(pdf);
                        listBoxOutput.Items.Add(pdf);
                    }
                    stream.Close();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void SearchFunction(string dir)
        {
            foreach (string subdir in Directory.GetDirectories(dir))
            {
                SearchFunction(subdir);
            }
            foreach (string fileName in Directory.GetFiles(dir))
            {
                string temp = fileName.Substring(fileName.LastIndexOf("\\"));
                if (temp.Contains(textToSearch)) //CSAK a fájl nevének vizsgálata, elérési útvonalé nem!
                {
                    listBoxOutput.Items.Add(fileName);
                }
                if (fileName.EndsWith(".doc")|fileName.EndsWith(".docx"))
                {
                    wordList.Add(fileName);
                }
                if (fileName.EndsWith(".xls") | fileName.EndsWith(".xlsx"))
                {
                    excelList.Add(fileName);
                }
                if (fileName.EndsWith(".pdf"))
                {
                    pdfList.Add(fileName);
                }
            }//foreach end

        }

        


        private void buttonSearch_Click_1(object sender, RoutedEventArgs e)
        {
            textToSearch = textBoxSearch.Text;
            listBoxOutput.Items.Clear();
            wordList.Clear();
            excelList.Clear();
            pdfList.Clear();
            SearchFunction(defLocation);

            double stopper = buttonSearch.ActualHeight;

            DocSearch();
            ExcelSearch();
            PdfSearch();
            //bw.RunWorkerAsync(defLocation);
            
        }
    }
}

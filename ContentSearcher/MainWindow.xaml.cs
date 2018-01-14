﻿using System;
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
     * Treeview rendszer
     * Gyökér ÉS/VAGY -FIX!
     *      -kifejezés -hozzáadandó
     *      -kifejezés
     *      Csoport ÉS/VAGY -hozzáadandó
     *          -kifejezés
     *          -kifejezés
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

        int lineCounter = 0; //hány filter sor van 
        int groupCounter = 0; //erre jobb megoldás fog kelleni!!!!!!

        public MainWindow()
        {
            InitializeComponent();

            
            bw.DoWork += Bw_DoWork;
            bw.RunWorkerCompleted += Bw_RunWorkerCompleted;

            ListFillUp(); //combobox-ok listájának feltöltése
            AddRootGroup(); //treeview gyökér item
        }

        private void ListFillUp()
        {
            logicList.Add("AND");
            logicList.Add("OR");

            subjectList.Add("FILE NAME");
            subjectList.Add("FILE CONTENT");

            operatorList.Add("EQUALS");
            operatorList.Add("NOT EQUALS");
            operatorList.Add("CONTAINS");
            operatorList.Add("NOT CONTAINS");
            operatorList.Add("STARTS WITH");
            operatorList.Add("NOT STARTS WITH");
            operatorList.Add("ENDS WITH");
            operatorList.Add("NOT ENDS WITH");
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

        private void AddRootGroup() //treeview item a kattintott treeview item alá! add line, add group, remove group
        {
            //kettő ilyen kell, egy remove-al és egy remove nélküli a gyökér group hozzáadásához
            TreeViewItem tvi = new TreeViewItem();
            ComboBox cb_logic = new ComboBox();
            Button bt_AddGroup = new Button();
            //Button bt_RemoveGroup = new Button();
            Button bt_AddLine = new Button();
            StackPanel sp = new StackPanel();
            Thickness thickness = new Thickness(2.5, 0, 2.5, 0);
            int buttonWidth = 35;
            int comboboxWidth = 65;

            cb_logic.ItemsSource = logicList;
            bt_AddGroup.Content = "[+]"; //vagy Add Group
            bt_AddLine.Content = "+"; //vagy Add Line
            //bt_RemoveGroup.Content = "Remove Group"; //vagy [-]
            bt_AddGroup.Margin = thickness;
            bt_AddLine.Margin = thickness;
            cb_logic.Margin = thickness;
            cb_logic.SelectedIndex = 0;
            cb_logic.Width = comboboxWidth;
            bt_AddGroup.Width = buttonWidth;
            bt_AddLine.Width = buttonWidth;

            bt_AddGroup.Click += Bt_AddGroup_Click;
            bt_AddLine.Click += Bt_AddLine_Click;
            
            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;
            

            tvi.Name = "rootTVI";
            this.RegisterName(tvi.Name, tvi);
            
            tvi.IsExpanded = true;
            sp.Children.Add(cb_logic);
            sp.Children.Add(bt_AddGroup);
            sp.Children.Add(bt_AddLine);
            tvi.Header = sp;
            //tvi.Items.Add(sp);
            treeViewFilter.Items.Add(tvi);
        }

        private void AddGroup(object source)
        {
            //kettő ilyen kell, egy remove-al és egy remove nélküli a gyökér group hozzáadásához
            //TreeViewItem tvi = new TreeViewItem();
            TreeViewItem sourceTVI = source as TreeViewItem;
            TreeViewItem tvi = new TreeViewItem();
            ComboBox cb_logic = new ComboBox();
            Button bt_AddGroup = new Button();
            Button bt_RemoveGroup = new Button();
            Button bt_AddLine = new Button();
            StackPanel sp = new StackPanel();
            Thickness thickness = new Thickness(2.5,0,2.5,0);
            int buttonWidth = 35;
            int comboboxWidth = 65;
            cb_logic.ItemsSource = logicList;
            bt_AddGroup.Content = "[+]"; //vagy Add Group
            bt_AddLine.Content = "+"; //vagy Add Line
            bt_RemoveGroup.Content = "[-]"; //vagy Remove Group
            bt_AddGroup.Margin = thickness;
            bt_AddLine.Margin = thickness;
            bt_RemoveGroup.Margin = thickness;
            cb_logic.Margin = thickness;
            cb_logic.SelectedIndex = 0;
            cb_logic.Width = comboboxWidth;
            bt_AddGroup.Width = buttonWidth;
            bt_AddLine.Width = buttonWidth;
            bt_RemoveGroup.Width = buttonWidth;

            bt_AddGroup.Click += Bt_AddGroup_Click;
            bt_AddLine.Click += Bt_AddLine_Click;


            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;
            tvi.IsExpanded = true;


            sp.Children.Add(cb_logic);
            sp.Children.Add(bt_AddGroup);
            sp.Children.Add(bt_AddLine);
            tvi.Header = sp;
            sourceTVI.Items.Add(tvi);
            //treeViewFilter.Items.Add(tvi);
        }

        private void Bt_AddGroup_Click(object sender, RoutedEventArgs e)
        {
            //sender kiderítése! és alatta lévő treeviewitem
            object childobject = sender;
            while (!(childobject is TreeViewItem))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }

            AddGroup(childobject);

        }

        private void Bt_AddLine_Click(object sender, RoutedEventArgs e)
        {
            //sender kiderítése! és alatta lévő treeviewitem
            object childobject = sender;
            while (!(childobject is TreeViewItem))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }

            AddLine(childobject);
        }

        private void Bt_delete_Click(object sender, RoutedEventArgs e)
        {
            object childobject = sender;
            while (!(childobject is StackPanel))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }
            StackPanel sp = childobject as StackPanel;
            //RemoveLine(childobject);
            MessageBox.Show(sp.ToString());
            childobject = sender;
            while (!(childobject is TreeViewItem))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }
            TreeViewItem parentTVI = childobject as TreeViewItem;
            MessageBox.Show(parentTVI.ToString());
            parentTVI.Items.Remove(sp);
            
        }

        private void AddLine(object source)
        {
            TreeViewItem sourceTVI = source as TreeViewItem;
            ComboBox cb_logic = new ComboBox();
            ComboBox cb_subject = new ComboBox();
            ComboBox cb_operator = new ComboBox();
            TextBox tb_expression = new TextBox();
            Button bt_delete = new Button();
            StackPanel sp = new StackPanel();
            Thickness thickness = new Thickness(2.5, 0, 2.5, 0);
            int logicWidth = 65;
            int subjectWidth = 110;
            int operatorWidth = 130;
            int expressionWidth = 250; //dinamikusnak kéne lennie?
            int buttonWidth = 35;

            cb_logic.ItemsSource = logicList;
            cb_subject.ItemsSource = subjectList;
            cb_operator.ItemsSource = operatorList;
            //cb_logic.Name = "comboBoxLogic_" + lineCounter.ToString();
            //cb_subject.Name = "comboBoxSubject_" + lineCounter.ToString();
            //cb_operator.Name = "comboBoxOperator_" + lineCounter.ToString();
            //tb_expression.Name = "textBoxExpression_" + lineCounter.ToString();
            //lineCounter++; //Sor számozás növelése
            //RegisterName(cb_logic.Name, cb_logic);
            //RegisterName(cb_subject.Name, cb_subject);
            //RegisterName(cb_operator.Name, cb_operator);
            //RegisterName(tb_expression.Name, tb_expression);


            cb_logic.SelectedIndex = 0;
            cb_subject.SelectedIndex = 0;
            cb_operator.SelectedIndex = 0;
            tb_expression.Text = "";
            bt_delete.Content = "X";

            cb_logic.Width = logicWidth;
            cb_operator.Width = operatorWidth;
            cb_subject.Width = subjectWidth;
            bt_delete.Width = buttonWidth;
            tb_expression.Width = expressionWidth;

            

            cb_logic.Margin = thickness;
            cb_operator.Margin = thickness;
            cb_subject.Margin = thickness;
            bt_delete.Margin = thickness;
            tb_expression.Margin = thickness;

            
            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;
            bt_delete.Click += Bt_delete_Click;

            sp.Children.Add(cb_logic);
            sp.Children.Add(cb_subject);
            sp.Children.Add(cb_operator);
            sp.Children.Add(tb_expression);
            sp.Children.Add(bt_delete);
            sourceTVI.Items.Add(sp);
            
        }

        private void RemoveLine(object source) //Adott sor törlése!!!!
        {
            StackPanel sourceSP = source as StackPanel;
            //ComboBox cb_logic = new ComboBox();
            //ComboBox cb_subject = new ComboBox();
            //ComboBox cb_operator = new ComboBox();
            //TextBox tb_expression = new TextBox();
            //Button bt_delete = new Button();
            //StackPanel sp = new StackPanel();
            //int logicWidth = 65;
            //int subjectWidth = 110;
            //int operatorWidth = 130;
            //int expressionWidth = 200; //dinamikusnak kéne lennie
            //int buttonWidth = 35;

            foreach (object item in sourceSP.Children)
            {
                //UnregisterName(item.Name as ComboBox)
            }

            object childobject = source;
            while (!(childobject is TreeViewItem))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }
            TreeViewItem parentTVI = childobject as TreeViewItem;
            parentTVI.Items.Remove(source);
            

        }


        private void buttonDEBUG_Add_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}

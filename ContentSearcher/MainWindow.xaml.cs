﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
    public partial class MainWindow : Window
    {
        
        string defLocation = @"E:\ZZZ_TESZTMAPPA";
        string searchRoot = string.Empty;

        List<string> fileList = new List<string>();

        List<string> outputList = new List<string>(); //végeredmény
        List<string> subresultList = new List<string>(); //köztes eredmény

        BackgroundWorker bw = new BackgroundWorker();

        List<object> logicList = new List<object>();
        List<object> subjectList = new List<object>();
        List<object> operatorList = new List<object>();

        int[] IDList = new int[200]; //nem végtelen, de bőven elég
                                    //használat: új TVI vagy StackPanel esetén elnevezni őket a legelső 'üres' tipusu tömb tag id-jével
                                    //0=üres, 1=TVI, 2=StackPanel
        

        public MainWindow()
        {
            InitializeComponent();

            
            bw.DoWork += Bw_DoWork;
            bw.RunWorkerCompleted += Bw_RunWorkerCompleted;

            ListFillUp(); //combobox-ok listájának feltöltése
            AddRootGroup(); //treeview gyökér item

            textBoxSearch.Text = defLocation; //DEBUG

        }

        private void ListFillUp()
        {
            logicList.Add("ÉS"); //0 AND
            logicList.Add("VAGY"); //1 OR
            
            subjectList.Add("Fájl tartalma"); //0 FILE CONTENT
            subjectList.Add("Fájl neve"); //1 FILE NAME

            operatorList.Add("Tartalmazza"); //0 CONTAINS
            operatorList.Add("Nem tartalmazza"); //1 NOT CONTAINS
            //operatorList.Add("Megegyezik"); //2 EQUALS
            //operatorList.Add("Nem egyezik meg"); //3 NOT EQUALS
            //operatorList.Add("Kezdődik vele"); //4 STARTS WITH
            //operatorList.Add("Nem kezdődik vele"); //5 NOT STARTS WITH
            //operatorList.Add("Véget ér vele"); //6 ENDS WITH
            //operatorList.Add("Nem ér véget vele"); //7 NOT ENDS WITH
        }

        private void Bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Bw_DoWork(object sender, DoWorkEventArgs e)
        {

            //Parallel.ForEach(Directory.GetFiles(e.Argument as string), );
        }

        private void WordSearch(string textToSearch, List<string> source)
        {
            try
            {
                Word.Application app = new Word.Application();
                foreach (string file in source)
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

        private void ExcelSearch(string textToSearch, List<string> source)
        {
            try
            {
                Excel.Application app = new Excel.Application();
                foreach (string xls in source)
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

        private void PdfSearch(string textToSearch, List<string> source)
        {
            try
            {
                foreach (string pdf in source)
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

        private void GetFileLists(string dir) // az összes felhasználható fájl listája pdf, word, excel
        {
            foreach (string subdir in Directory.GetDirectories(dir))
            {
                GetFileLists(subdir);
            }
            foreach (string fileName in Directory.GetFiles(dir))
            {
                if (fileName.EndsWith(".doc")|fileName.EndsWith(".docx"))
                {
                    fileList.Add(fileName);
                }
                if (fileName.EndsWith(".xls") | fileName.EndsWith(".xlsx"))
                {
                    fileList.Add(fileName);
                }
                if (fileName.EndsWith(".pdf"))
                {
                    fileList.Add(fileName);
                }
            }//foreach end

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

            bt_AddGroup.Background = Brushes.LawnGreen;
            bt_AddLine.Background = Brushes.LightGreen;

            bt_AddGroup.FontWeight = FontWeights.Bold;
            bt_AddLine.FontWeight = FontWeights.Bold;

            bt_AddGroup.Click += Bt_AddGroup_Click;
            bt_AddLine.Click += Bt_AddLine_Click;
            
            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;

            //NAMING
            for (int i = 0; i < IDList.Length; i++)
            {
                if (IDList[i] == 0)
                {
                    IDList[i] = 1;
                    tvi.Name = "ID_" + i.ToString();
                    
                    break;
                }
            }
            tvi.IsExpanded = true;
            sp.Children.Add(cb_logic);
            sp.Children.Add(bt_AddGroup);
            sp.Children.Add(bt_AddLine);
            tvi.Header = sp;
            //tvi.Items.Add(sp);
            treeViewFilter.Items.Add(tvi);
            RegisterName(tvi.Name, tvi);
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

            bt_AddGroup.FontWeight = FontWeights.Bold;
            bt_AddLine.FontWeight = FontWeights.Bold;
            bt_RemoveGroup.FontWeight = FontWeights.Bold;


            bt_AddGroup.Background = Brushes.LawnGreen;
            bt_AddLine.Background = Brushes.LightGreen;
            bt_RemoveGroup.Background = Brushes.OrangeRed;
            
            //NAMING
            for (int i = 0; i < IDList.Length; i++)
            {
                if (IDList[i] == 0)
                {
                    IDList[i] = 1;
                    tvi.Name = "ID_" + i.ToString();
                    
                    break;
                }
            }
            bt_AddGroup.Click += Bt_AddGroup_Click;
            bt_AddLine.Click += Bt_AddLine_Click;
            bt_RemoveGroup.Click += Bt_RemoveGroup_Click;

            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;
            tvi.IsExpanded = true;
          
            sp.Children.Add(cb_logic);
            sp.Children.Add(bt_AddGroup);
            sp.Children.Add(bt_AddLine);
            sp.Children.Add(bt_RemoveGroup);
            tvi.Header = sp;
            sourceTVI.Items.Add(tvi);
            RegisterName(tvi.Name, tvi);
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

        private void AddLine(object source)
        {
            TreeViewItem sourceTVI = source as TreeViewItem;
            //ComboBox cb_logic = new ComboBox();
            ComboBox cb_subject = new ComboBox();
            ComboBox cb_operator = new ComboBox();
            TextBox tb_expression = new TextBox();
            Button bt_delete = new Button();
            StackPanel sp = new StackPanel();
            Thickness thickness = new Thickness(2.5, 0, 2.5, 0);
            //int logicWidth = 65;
            int subjectWidth = 110;
            int operatorWidth = 130;
            int expressionWidth = 250; //dinamikusnak kéne lennie?
            int buttonWidth = 35;

            //cb_logic.ItemsSource = logicList;
            cb_subject.ItemsSource = subjectList;
            cb_operator.ItemsSource = operatorList;
            
            //cb_logic.SelectedIndex = 0;
            cb_subject.SelectedIndex = 0;
            cb_operator.SelectedIndex = 0;
            tb_expression.Text = "";
            bt_delete.Content = "X";

            //cb_logic.Width = logicWidth;
            cb_operator.Width = operatorWidth;
            cb_subject.Width = subjectWidth;
            bt_delete.Width = buttonWidth;
            tb_expression.Width = expressionWidth;

            //cb_logic.Margin = thickness;
            cb_operator.Margin = thickness;
            cb_subject.Margin = thickness;
            bt_delete.Margin = thickness;
            tb_expression.Margin = thickness;
            
            sp.FlowDirection = FlowDirection.LeftToRight;
            sp.Orientation = Orientation.Horizontal;

            bt_delete.Click += Bt_delete_Click;
            bt_delete.Background = Brushes.LightPink;
            bt_delete.FontWeight = FontWeights.Bold;

            //NAMING
            for (int i = 0; i < IDList.Length; i++)
            {
                if (IDList[i] == 0)
                {
                    IDList[i] = 2;
                    sp.Name = "ID_" + i.ToString();
                    
                    break;
                }
            }

            //sp.Children.Add(cb_logic);
            sp.Children.Add(cb_subject);
            sp.Children.Add(cb_operator);
            sp.Children.Add(tb_expression);
            sp.Children.Add(bt_delete);
            sourceTVI.Items.Add(sp);
            RegisterName(sp.Name, sp);
        }

        private void Bt_delete_Click(object sender, RoutedEventArgs e)
        {
            object childobject = sender;
            bool conti = false;
            while (!(childobject is StackPanel))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
            }
            StackPanel SPToDelete = childobject as StackPanel; //ez még stimmel!
            //MessageBox.Show(SPToDelete.Name);
            //IDList 1-esein végigmenni, melyik TVI
            childobject = sender;
            while (!(childobject is TreeViewItem & conti))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
                if (childobject is TreeViewItem)
                {
                    TreeViewItem tempobj = childobject as TreeViewItem;
                    if (tempobj.Name != "") //üres TVI (WTF?!) átugrása
                    {
                        conti = true;
                    }
                }
            }
            TreeViewItem parentTVI = childobject as TreeViewItem;
            string temp = SPToDelete.Name;
            int index = Convert.ToInt32(temp.Substring(temp.IndexOf("_")+1));
            IDList[index] = 0; //nullázás törlés miatt
            parentTVI.Items.Remove(SPToDelete);
            UnregisterName(temp);
        }

        private void Bt_RemoveGroup_Click(object sender, RoutedEventArgs e)
        {
            //get parent, remove tvi and children, unregister names
            // delete group
            object childobject = sender;
            bool conti = false;
            while (!(childobject is TreeViewItem & conti))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
                if (childobject is TreeViewItem)
                {
                    //TreeViewItem tempTVI = childobject as TreeViewItem;
                    conti = true;
                }
            }
            TreeViewItem sourceTVI = childobject as TreeViewItem; //törlendő tvi
            conti = false; //folytatás tovább
            while (!(childobject is TreeViewItem & conti))
            {
                childobject = VisualTreeHelper.GetParent(childobject as DependencyObject);
                if (childobject is TreeViewItem)
                {
                    TreeViewItem tempTVI = childobject as TreeViewItem;
                    conti = true;

                }
            }
            TreeViewItem parentTVI = childobject as TreeViewItem;

            UnRegTVIItems(sourceTVI);

            string temp = sourceTVI.Name;
            parentTVI.Items.Remove(sourceTVI);
            UnregisterName(temp);
        }

        private void UnRegTVIItems(TreeViewItem source)
        {
            foreach (object item in source.Items)
            {
                if (item is TreeViewItem)
                {
                    TreeViewItem itemTVI = item as TreeViewItem;
                    UnRegTVIItems(itemTVI); //rekurziv
                    UnregisterName(itemTVI.Name);
                    int index = Convert.ToInt32(itemTVI.Name.Substring(itemTVI.Name.IndexOf("_") + 1));
                    IDList[index] = 0;
                }
                if (item is StackPanel)
                {
                    StackPanel itemSP = item as StackPanel;
                    UnregisterName(itemSP.Name);
                    int index = Convert.ToInt32(itemSP.Name.Substring(itemSP.Name.IndexOf("_") + 1));
                    IDList[index] = 0;
                }
            }
        }

        private void buttonBrowse_Click(object sender, RoutedEventArgs e)
        {
            Ookii.Dialogs.Wpf.VistaFolderBrowserDialog dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            if (dialog.ShowDialog(this).GetValueOrDefault())
            {
                textBoxSearch.Text = dialog.SelectedPath;
            }
        }

        private void textBoxSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (Directory.Exists(textBoxSearch.Text))
            {
                buttonSearch.IsEnabled = true;
            }
            else
            {
                buttonSearch.IsEnabled = false;
            }
        }


        private void buttonSearch_Click_1(object sender, RoutedEventArgs e)
        {
            searchRoot = textBoxSearch.Text; //textbox mező beolvasása

            listBoxOutput.Items.Clear(); //takarítás
            fileList.Clear();
            outputList.Clear();

            GetFileLists(searchRoot); //fileList feltöltése
            /////////////////////////////////////////////////
            TreeViewItem root = FindName("ID_0") as TreeViewItem; //mindig ez a 0-ás!
            sajtkukac(root);
            //backgroundworker-nek hogyan lehetne odaadni a treeview-t? object copy, egyéb?
       
        }

        private void sajtkukac(TreeViewItem rootTVI)
        {
            //Get controls
            //If contains TVI -> start again from here (nested)
            //Else - get expression
            //combine results into one big expression
            //
            //return result

            foreach (object item in rootTVI.Items)
            {
                if (item is TreeViewItem)//beágyazott csoport
                {
                    TreeViewItem childTVI = item as TreeViewItem;
                    sajtkukac(childTVI);
                }//TVI vége
                if (item is StackPanel)//tartalmazott kifejezések
                {
                    StackPanel childSP = item as StackPanel;
                    foreach (object spItem in childSP.Children)
                    {
                        if (spItem is ComboBox)
                        {
                            ComboBox cb = spItem as ComboBox;
                            switch (cb.Text)
                            {
                                //------------------Forrás
                                case "Fájl neve":

                                    break;
                                case "Fájl tartalma":

                                    break;
                                //-------------------Forrás vége
                                //-------------------Operátor
                                case "Tartalmazza":

                                    break;
                                case "Nem tartalmazza":

                                    break;
                                //case "Megegyezik":

                                //    break;
                                //case "Nem egyezik meg":

                                //    break;
                                //case "Kezdődik vele":

                                //    break;
                                //case "Nem kezdődik vele":

                                //    break;
                                //case "Véget ér vele":

                                //    break;
                                //case "Nem ér véget vele":

                                //    break;
                                //-------------------Operátor vége
                                default:
                                    break;
                            }
                        }
                        if (spItem is TextBox)
                        {
                            TextBox tb = spItem as TextBox;

                        }
                    }//stackpanel elemeinek vége

                }//stackpanel vége
            }
        }


        

    }
}
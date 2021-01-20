using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
//using System.Windows.Forms;
//using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Diagnostics;

namespace ExcelReadingApp
{
    public partial class Form1 : Form
    {
        public string strFilename1;
        public string FileInputDir = @""; //the file location changer meanwhile debugging
        public string File1FullPath,
            File2FullPath = @"",
            File1Name ,//= "AlsoEnergy(_Single_)_latestFile2.xls",
            File2Name = "",
            File1NameTrimmed;

        public List<int> ColumnNumberToAddFromFile2 = new List<int>();
        public List<string> ColumnNameToAddFromFile2 = new List<string>();
        public List<int> ColumnNumberToDeleteFromFile1 = new List<int>();
        public List<string> ColumnNameToDeleteFromFile1 = new List<string>();

        public const int APPEND = 1, NEWLine = 2;

        RootDirectoriesExplorer RE = new RootDirectoriesExplorer();
        Excel_MS EX = new Excel_MS();
        ExcelLoad EL = new ExcelLoad();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //RootDirectoriesExplorer RE = new RootDirectoriesExplorer();
            //Excel_MS EX = new Excel_MS();
            //ExcelLoad EL = new ExcelLoad();
            EL.xmlLoadData();
            RE.DirectoriesExplorer();
            EX.XLSExtraction(RE.FileNames, RE.FileDirecrtory, RE.DirNames);

        //    //foreach(string s in RE.DirNames)
        //    //{
        //    //    //DisplayText(s, APPEND);
        //    //    Excel_MS Ex = new Excel_MS(LatestFileSort(), 1);//LatestFileSort()
        //    //    DisplayText(File1FullPath, APPEND);
        //    //    Ex.ReadCell(1);
        //    //}
        //    //Excel_MS Ex = new Excel_MS(File1FullPath, 1);//LatestFileSort()
        //    //DisplayText(File1FullPath, APPEND);
        //    //Ex.ReadCell(1);
        //    Excel_MS Ex1 = new Excel_MS(File2FullPath, 1);//ReferencefileSort()//changed 
        //    DisplayText(File2FullPath, APPEND);//File2FullPath
        //    Ex1.ReadCell(1);
        //    //XlsCompareToAdd(Ex.Dataset, Ex1.Dataset);
            
        //    Excel_MS Ex2 = new Excel_MS(File1FullPath, 1);
        //    Ex2.CheckFortheEmptyCells(ColumnNumberToDeleteFromFile1,ColumnNameToDeleteFromFile1);
        //    DisplayText("Columns SAFE!", APPEND);
        //    for (int counter = 0; counter < ColumnNameToDeleteFromFile1.Count; counter++) { DisplayText(ColumnNameToDeleteFromFile1[counter], APPEND); }
        //    Ex2.DeleteCells(ColumnNumberToDeleteFromFile1, ColumnNameToDeleteFromFile1, File1Name);
        //    Ex2.AddCells(ColumnNumberToAddFromFile2, ColumnNameToAddFromFile2, File1Name);
        //    Excel_MS Ex3 = new Excel_MS(@"G:\_ShipmentFiles\vishalModified\" + File1Name, 1);
        //    Ex3.CleanNamelessColumns(ColumnNumberToDeleteFromFile1);
        //    //if (ColumnsToKeep != Ex.Dataset.Count)
        //    //    DisplayText("Good File. Does Not need Modification.",APPEND);

        //    //Ex.DeleteCells();
        //    //Ex.AddCells(ColumnNumberToAddFromFile2, ColumnNameToAddFromFile2);
        //    //Ex.AddCells(ColumnNumberToAddFromFile2, ColumnNameToAddFromFile2, File1Name);

        }
        #region Adjoining Functions
        private void Form1_Load(object sender, EventArgs e)
        {
            //OpenFile();
            //DirectoryInfo[] dir = new DirectoryInfo(@"G:\").GetDirectories("*.*", SearchOption.TopDirectoryOnly);// AllDirectories
            //foreach (DirectoryInfo d in dir)
            //{
            //    comboBox1.Items.Add(d.Name);
            //}
        }

        public void OpenFile()
        {
            //Excel_MS Ex = new Excel_MS(@"G:\demo_files\DemoBook.xlsx",1);
            //MessageBox.Show(Ex.ReadCell(0, 0));

        }

        public string CurrentDateAndTime()
        {
            DateTime lastupdated = DateTime.Today;
            string dateFormatYY_MM_dd = lastupdated.ToString("yyyy_MM_dd");
            return dateFormatYY_MM_dd;
        }

        public string LatestFileSort()
        {
            DateTime lastupdated = DateTime.MinValue;
            strFilename1 = FileInputDir;   
            string Folder = strFilename1;
            
            var files = new DirectoryInfo(Folder).GetFiles("*.xls");

            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime > lastupdated)
                {
                    lastupdated = file.LastWriteTime;
                    File1FullPath = file.FullName;
                    File1Name = file.Name;
                }
            }
            int temp_index = File1Name.IndexOf("_");
            File1NameTrimmed = File1Name.Substring(temp_index+1, 15);
            int temp_index1 = File1NameTrimmed.IndexOf("_");
            temp_index = temp_index + temp_index1+1;//additional 1 char helps to get the complete last word
            File1NameTrimmed = File1Name.Substring(0, temp_index);
            return File1FullPath;
        }

        public string ReferencefileSort()
        {
            DateTime lastupdated = DateTime.Today;
            strFilename1 = FileInputDir;
            string Folder = strFilename1;

            var files = new DirectoryInfo(Folder).GetFiles("*.xls");

            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime < lastupdated)
                {
                    if(file.Name.Contains(File1NameTrimmed)) //softcode it
                    {
                        lastupdated = file.LastWriteTime;
                        File2FullPath = file.FullName;
                        File2Name = file.Name;
                    }
                }
            }
            return File2FullPath;
        }

        private void XlsCompareToAdd(List<string> TemoList1, List<string> TemoList2)
        {
            for (int reference1 =0,reference2=0;reference1 < TemoList2.Count ||reference2 < TemoList1.Count; reference1++,reference2++)//int reference = TemoList1.Count - 1; reference >= 0; reference--
            {
                if(reference1< TemoList2.Count)
                {
                    if (!TemoList1.Contains(TemoList2[reference1]))
                    {
                        ColumnNumberToAddFromFile2.Add(reference1 + 1);
                        ColumnNameToAddFromFile2.Add(TemoList2[reference1]);
                    }
                }

                if(reference2 < TemoList1.Count)//checks the condition
                {
                    if (!TemoList2.Contains(TemoList1[reference2]))
                    {
                        ColumnNumberToDeleteFromFile1.Add(reference2+1);
                        ColumnNameToDeleteFromFile1.Add(TemoList1[reference2]);
                        //if (TemoList1[reference] == TemoList2[reference])
                        //{

                        //}
                        //ColumnNumberToDeleteFromFile1.Add(reference);
                        //ColumnNameToDeleteFromFile1.Add(TemoList2[reference])
                    }
                }
                ////int numOfDuplicates = 1;
                //for (int comparingTo = TemoList2.Count - 2; comparingTo >= 0; comparingTo--)
                //{
                //    if (TemoList1[reference] == TemoList2[comparingTo])
                //        ColumnsToKeep++;

                //    else if(TemoList1[reference]!=TemoList2[comparingTo])
                //        ColumnNumberToDelete.Add(reference);
                //        ColumnNameToDelete.Add(TemoList2[reference]);
                //}
            }

        }

        public void DisplayText(string textContent,int Condition)
        {
            switch(Condition)
            {
                case APPEND:
                    richTextBox1.AppendText("\r\n" + textContent);
                    break;
                case NEWLine:
                    richTextBox1.Text = textContent;
                    break;

                default:
                    richTextBox1.Text = textContent;
                    break;
            }

            //if(Condition.Contains("Append"))
            //{
            //    richTextBox1.AppendText("\r\n"+textContent);
            //}
            //else if(Condition.Contains("New"))
            //{
            //    richTextBox1.Text = textContent;
            //}
            //else
            //{
            //    richTextBox1.Text = textContent;
            //}
        }
        #endregion Adjoining Functions
    }
}

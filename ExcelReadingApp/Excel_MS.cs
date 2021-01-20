using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelReadingApp
{
    class Excel_MS
    {
        _Application excel = new _Excel.Application();

        Workbook wb;
        Worksheet ws;
        Form1 F1 = new Form1();

        public List<string> Dataset = new List<string>();
        public List<string> DatasetRow2 = new List<string>();
        public List<string> RemovedColumns = new List<string>();
        string path;
        public Excel_MS()//string path,int sheet
        {
            //this.path = path;
            //wb = excel.Workbooks.Open(path);
            //ws = wb.Worksheets[sheet];//(Worksheet)
        }
        ~Excel_MS()//string path, int sheet
        { }

        public void ReadCell(int sheetRow)
        {
            int sheetColumn = 1; 
            //int IndexString =0;
            string fileName = @"c:\Temp\ExcelLayout.xml";
            while (ws.Cells[sheetRow,sheetColumn].Value !=null)
            {
                Dataset.Add(ws.Cells[sheetRow, sheetColumn].Value2);
                if (ws.Cells[sheetRow + 1, sheetColumn].Value2 != null)
                {
                    dynamic demo = ws.Cells[2, sheetColumn].Value2;
                    DatasetRow2.Add(demo+string.Empty);
                }
                else
                    DatasetRow2.Add(string.Empty);
                sheetColumn++;
            }
            ExcelLayout layout = ExcelLayoutManager.Initialize(Dataset,DatasetRow2);
            ExcelLayoutManager.Save(fileName, layout, true);
        }
        public void AddCells(List<int> temp1, List<string> temp2,string File1Name)
        {
            //ColumnNumberToAddFromFile2
            //ColumnNameToAddFromFile2
            int counter = 0;
            try
            {
                while (counter<temp1.Count)
                {
                    ws.Columns[temp1[counter]].Cells.Insert();//ws.Columns[temp1[counter]+counter].Cells.Insert();
                    ws.Cells[1,temp1[counter]].Value2 = temp2[counter]+"."+counter; //this works as well// this does not insert the column into the sheet it just replaces the name it has previously.//ws.Cells[1,temp1[counter] + counter].Value2 = "Cell"+ temp1[counter]+"."+counter; 
                    //ws.Columns.
                    //ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
                    ////wb.Save();
                    ////wb.Close();
                    //excel.Workbooks.Close();
                    counter++;
                }
                ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
                wb.Close();
                excel.Workbooks.Close();
                //ws.Columns["D"].Clear();//Works
                //ws.Columns["D"].Delete();//works

                //ws.Columns[2].Cells.Insert();
                //ws.Cells[2,2].Value2 = "Cell B2."; //this works as well// this does not insert the column into the sheet it just replaces the name it has previously.
                //ws.Rows.Cells[1,1].Value2 = "column B.";//works-> row column
                ////ws.Columns.Insert(3);
                ////ws.Columns[demo].Insert();
                //ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
                //wb.Save();
                //wb.Close();
                //excel.Workbooks.Close();
            }
            catch(Exception ex){ MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error); }
 
        }
        public void DeleteCells(List<int> temp1, List<string>temp2, string File1Name)
        {
            int counter = 0;
            while (counter<temp1.Count)
            {
                ws.Columns[temp1[counter]+1].Clear();//Works
                //ws.Columns[temp1[counter]+1].Delete();//works
                counter++;
            }
            ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
            //excel.Workbooks.Close();
            //int counter = temp1.Count-1;
            //while (counter > 0)
            //{
            //    ws.Columns[temp1[counter]].Add();
            //}
            //ws.Cells.EntireColumn.Delete(1);
            //ws.Columns[3].Clear();
            //ws.Columns[3].Delete();
            //wb.Save(); wb.Close();

        }

        public void CheckFortheEmptyCells(List<int> CNumberDFF1, List<string> CNameDFF1)
        {
            //ColumnNumberToDeleteFromFile1  CNumberDFF1
            //ColumnNameToDeleteFromFile1 CNameDFF1
            for(int counter=0;counter< CNumberDFF1.Count;)
            {
                //can be implemented
                if (ws.Cells[2, CNumberDFF1[counter]].Value2 != null)
                {
                    CNumberDFF1.RemoveAt(counter);
                    CNameDFF1.RemoveAt(counter);
                    if (counter != 0) { counter--; }
                    else { counter = 0; }
                }
                counter++;
            }
        }

        public void CleanNamelessColumns(List<int> CNumberDFF1)
        {
            int FlagtoExit = 0;
            for (int counter = 1; counter < 200;)
            {
                //can be implemented
                if (ws.Cells[1, counter].Value2 == null)
                {
                    ws.Columns[counter].Delete();
                    if (counter != 0) { counter--; FlagtoExit++; }
                    else { counter = 0; }
                }
                else { FlagtoExit=0; }
                counter++;
                if (FlagtoExit >= 10)//FlagtoExit >= 10//FlagtoExit >= CNumberDFF1.Count
                    break;
            }
            wb.Save(); excel.Workbooks.Close();
        }

        public void XLSExtraction(List<string>InputFilename,List<string>FileDir,List<string>DirectoryName)
        {
           
            for(int counter =0;counter< FileDir.Count; counter++)// FileDir.Count
            {
                wb = excel.Workbooks.Open(FileDir[counter]);
                ws = wb.Worksheets[1];//sheet

                int sheetColumn = 1;
                int sheetRow = 1;
                string trimAddress = InputFilename[counter].Substring(0,InputFilename[counter].Length-3);
                //string fileName = "F"+trimAddress;
                Directory.CreateDirectory(@"F:\ShipmentsXMLfiles\" + DirectoryName[counter]);
                string OutputFileName = @"F:\ShipmentsXMLfiles\" + DirectoryName[counter] + "\\" + trimAddress+"xml"; //chjange the location
                while (ws.Cells[sheetRow, sheetColumn].Value != null)
                {
                    try
                    {
                        dynamic TempDynamicString1 = ws.Cells[sheetRow, sheetColumn].Value2;
                        Dataset.Add(TempDynamicString1 + string.Empty);
                        if (ws.Cells[sheetRow + 1, sheetColumn].Value2 != null)
                        {
                            dynamic TempDynamicString2 = ws.Cells[2, sheetColumn].Value2;
                            DatasetRow2.Add(TempDynamicString2 + string.Empty);
                        }
                        else
                            DatasetRow2.Add(string.Empty);
                        sheetColumn++;
                    }
                    catch
                    {
                        sheetColumn++;
                    }
                    
                }
                ExcelLayout layout = ExcelLayoutManager.Initialize(Dataset, DatasetRow2);
                ExcelLayoutManager.Save(OutputFileName, layout, true);
                wb.Close();
                Dataset.Clear();
                DatasetRow2.Clear();
            }
        }

        /*       bool result = Dataset.Contains("FirmwareRevision");
            if(result)
            {
                IndexString = Dataset.IndexOf("FirmwareRevision");
                IndexString++;
                ws.Columns[IndexString].Clear();
                RemovedColumns.Add("" + IndexString);
                //ws.Cells.EntireColumn.Delete(IndexString);
                //ws.SaveAs(@"G:\demo_files\DemoAutoSavedFile.xls");
                wb.Save();
                wb.Close();


            ws.Columns[3].Clear();   //the following 4 steps works fine.
            ws.Columns[3].Delete();
            wb.Save(); wb.Close();

            ws.Columns["D"].Insert(10); //here D is the column in excel file

            ws.Columns["D"].Clear();
            ws.Columns["D"].Delete();
            ws.Cells[2,2].Value2 = "Cell B2.";          //this works as well
            ws.Rows.Cells[1,1].Value2 = "column B.";        //works-> row column

        while loop from

         ws.Columns[counter].Delete();
                    ws.Columns.Insert(counter);//temp1[counter]+counter
                    //ws.Rows.Insert(0,3);  //[temp1[counter]+counter].Text(""+temp2[counter]);
                    //ws.Rows.Insert(temp2[counter]);
                    //ws.Columns.Insert(3);
                    //ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\"+ File1Name);
                    //while (counter <2)// temp1.Count)
                    //{
                    //    //ws.Columns[temp1[counter]].Add(temp2[counter]);
                    //    ws.Columns["C"].Insert();//temp1[counter]  //"\""+temp2[counter]+ "\""



        delete loop
                        //if(ws.Cells[2, CNumberDFF1[counter]].Value2 == null)
                //{
                //    counter++;
                //}
                //else
                //{
                //    CNumberDFF1.RemoveAt(counter);
                //    CNameDFF1.RemoveAt(counter);
                //    if (counter != 0) { counter--; }
                //    else { counter = 0; }
                //}


            }
         
         */
    }
}

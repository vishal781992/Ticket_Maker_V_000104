using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using _Excel = Microsoft.Office.Interop.Excel;//excel workbooks are using this
using excel1 = Microsoft.Office.Interop.Excel;
using ExcelSpecial = Microsoft.Office.Interop.Excel;
using System.Xml;

namespace ExcelReadingApp
{
    class ExcelProcessor
    {
        _Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        _Application excel = new _Excel.Application();
        _Application excelSpcl = new _Excel.Application();

        public int counterForFileGeneratedInXml = 0;
        //Workbook wb;
        //Worksheet ws;
        Form1 F1 = new Form1();

        public List<string> Dataset = new List<string>();
        public List<string> DatasetRow2 = new List<string>();
        public List<string> RemovedColumns = new List<string>();
        public List<int> ColumnsToRemoveFromNewExcelSheet = new List<int>();

        public List<string> DatDBColumnNames = new List<string>();
        public List<string> FileColumnNames = new List<string>();
        public List<string> ValueForColumnStatics = new List<string>();
        public List<string> MergeEvents = new List<string>();

        public List<string> Place1 = new List<string>();
        public List<string> Place2 = new List<string>();
        public List<string> Place3 = new List<string>();
        public List<string> Place4 = new List<string>();

        bool flag_XML_PLACE = true;


        public ExcelProcessor() { }

        #region WriteNewXLS
        /*Function working: 
         * This function works towards creating new excel sheet from the data supplied from the function call in the main button.Tab1
         * List<string> ColumnValue - is the column names displayedin the excel in row 1,
         * int RowCounter - this is used to count how long the excel sheet is and to run the for loops.
         * dynamic[,] ArrayMessageFromDatabase - Contains all the database or in other words all the data gathered from the test query.
         * List<string> SimCardIDCode - not used in the function although called (maybe deleted soon),
         * [Optional] string Path - its the path of the file, The Parent folder is usually _Shipment or Vishal_Shipments. This identifies the company name and makes the path dynamically.
         */
        public void WriteANewExcel(List<string> ColumnValue, int RowCounter, dynamic[,] ArrayMessageFromDatabase, [Optional] string Path)//, List<string> RowValue,string Path
        {
            //path is null should be added

            _Excel.Workbook xlWorkBook;
            _Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            //xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkBook = excel.Workbooks.Add(misValue);
            xlWorkSheet = (_Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //XMLParser x = new XMLParser();

            for (int column = 0; column < ColumnValue.Count; column++)
            {
                xlWorkSheet.Columns[column + 1].NumberFormat = "@";
            }

            for (int row = 0; row <= RowCounter; row++)
            {
                for (int column = 0; column < ColumnValue.Count; column++)
                {
                    try
                    {
                        if (row == 0)
                        {
                            if (ColumnValue[column].Contains("CommID1") || ColumnValue[column].Contains("CommID2") || ColumnValue[column].Contains("CommID3") || ColumnValue[column].Contains("CommID4"))
                            {
                                xlWorkSheet.Cells[row + 1, column + 1] = "CommID";
                                xlWorkSheet.Cells[row + 1, column + 1].Font.Bold = true;
                                xlWorkSheet.Rows[1].Interior.Color = System.Drawing.Color.Gray;
                            }
                            else
                            {
                                xlWorkSheet.Cells[row + 1, column + 1] = ColumnValue[column];
                                xlWorkSheet.Cells[row + 1, column + 1].Font.Bold = true;
                                xlWorkSheet.Rows[1].Interior.Color = System.Drawing.Color.Gray;
                            }
                        }

                        else
                        {
                            if (!ColumnValue[column].Contains("IMEI") && !ColumnValue[column].Contains("SimCardID") && ArrayMessageFromDatabase[row - 1, column].Length >= 13)
                            {
                                xlWorkSheet.Cells[row + 1, column + 1] = "\'" + ArrayMessageFromDatabase[row - 1, column];
                            }
                            else if (ColumnValue[column].Contains("KwhUsage"))
                            {
                                string tempFormated = string.Format("{0:00000}", ArrayMessageFromDatabase[row - 1, column]);
                                xlWorkSheet.Cells[row + 1, column + 1] = tempFormated;
                            }
                            else
                            {
                                xlWorkSheet.Cells[row + 1, column + 1] = ArrayMessageFromDatabase[row - 1, column];
                            }

                        }

                    }
                    catch { }
                }
            }

            try
            { xlWorkBook.SaveAs(Path); }//@"C:\temp\demoFile.xlsx"
            catch { MessageBox.Show("The Excel File is not being Saved!", "Warning>> ", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            //xlWorkBook.Close(true, misValue, misValue);
            xlWorkBook.Close(0);

            //excel.Quit();
            //Marshal.ReleaseComObject(xlWorkSheet);
            //Marshal.ReleaseComObject(xlWorkBook);
            //Marshal.ReleaseComObject(xlApp);
            //Marshal.ReleaseComObject(excel);
            // Get rid of everything - close Excel

            while (Marshal.ReleaseComObject(xlWorkBook) > 0) { }
            xlWorkBook = null;
            while (Marshal.ReleaseComObject(xlWorkSheet) > 0) { }
            xlWorkSheet = null;
            GC();
            excel.Quit();
            xlApp.Quit();
            while (Marshal.ReleaseComObject(excel) > 0) { }
            excel = null;
            while (Marshal.ReleaseComObject(xlApp) > 0) { }
            excel = null;
            GC();

            //Marshal.ReleaseComObject(wb);
        }
        public static void GC()
        {
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        #endregion WriteNewXLS

        #region ExcelExtraction
        public int ExcelExtraction(List<string> InputFilename, List<string> FileDir, List<string> DirectoryName, string ParentFolderToTakeFrom, String ParentFolderToStickTo)
        {
            string outputDir;
            string OutputFileName;
            _Excel.Workbook wb;
            _Excel.Worksheet ws;

            for (int counter = 0; counter < FileDir.Count; counter++)// FileDir.Count
            {
                outputDir = string.Empty;
                OutputFileName = string.Empty;
                if (!string.Equals(InputFilename[counter], "NoFileFound"))//NoFileFound InputFilename[counter]
                {

                    wb = excel.Workbooks.Open(FileDir[counter]);
                    ws = wb.Worksheets[1];//sheet

                    int sheetColumn = 1;
                    int sheetRow = 1;
                    string trimAddress = InputFilename[counter].Substring(0, InputFilename[counter].Length - 3);

                    if (Directory.Exists(ParentFolderToTakeFrom + "\\" + DirectoryName[counter]))
                    {
                        //Directory.CreateDirectory(ParentFolderToTakeFrom + DirectoryName[counter]);
                        outputDir = ParentFolderToStickTo + DirectoryName[counter];
                        OutputFileName = ParentFolderToStickTo + DirectoryName[counter] + "\\" + trimAddress + "xml"; //chjange the location
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
                        //ExcelLayoutManager.Save(OutputFileName, layout, true);
                        //wb.Close();

                        wb.Close(0);

                        #region Commented
                        //while (Marshal.ReleaseComObject(wb) > 0) { }
                        //wb = null;
                        //while (Marshal.ReleaseComObject(ws) > 0) { }
                        //ws = null;
                        //GC();
                        //excel.Quit();
                        //xlApp.Quit();
                        //while (Marshal.ReleaseComObject(excel) > 0) { }
                        //excel = null;
                        //while (Marshal.ReleaseComObject(xlApp) > 0) { }
                        //excel = null;
                        //GC();
                        #endregion Commented
                    }
                    XMLParser XML0 = new XMLParser();
                    //XMLCreatorFromExcel(string OutputFileCompletepath,string outputDir, List<string> Dataset, List<string> DatasetRow2)
                    counterForFileGeneratedInXml += XML0.XMLCreatorFromExcel(OutputFileName, outputDir, Dataset, DatasetRow2);
                    Dataset.Clear();
                    DatasetRow2.Clear();
                }

            }
            return counterForFileGeneratedInXml;

        }
        #endregion ExcelExtraction

        #region ExcelModifierTemp

        public int ExcelModifierFunction(string pathForFile, int ColumnCount)
        {
            try
            {
                excel1.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                _Application excel = new excel1.Application();

                excel1.Workbook xlWorkBook;
                excel1.Worksheet xlsWorkSheet;

                xlWorkBook = excel.Workbooks.Open(pathForFile);
                xlsWorkSheet = xlWorkBook.Worksheets[1];//sheet

                int ColumnNumber = 1;
                /*
                 * Note: row 1 is titles , So the program checkas the 2nd row for the values, if the 2nd row is empty it assumes that the column is empty.
                 * which is true in most of the cases.
                 */
                //while (xlsWorkSheet.Cells[2, ColumnNumber].Value != null && ColumnNumber < (ColumnCount))
                while ((xlsWorkSheet.Cells[1, ColumnNumber].Value2 != null || !string.IsNullOrEmpty(xlsWorkSheet.Cells[2, ColumnNumber].Value2)) && ColumnNumber < (ColumnCount))
                {
                    dynamic TestPointString = xlsWorkSheet.Cells[2, ColumnNumber].Value2;
                    dynamic debugString = xlsWorkSheet.Cells[1, ColumnNumber].Value;
                    if (string.IsNullOrEmpty(xlsWorkSheet.Cells[2, ColumnNumber].Value))
                    {
                        if (xlsWorkSheet.Cells[1, ColumnNumber].Value == "Comments") { break; }
                        else { xlsWorkSheet.Columns[ColumnNumber].Delete(); ColumnCount -= 1; }
                    }
                    else { ColumnNumber++; }
                }

                xlWorkBook.Save(); xlWorkBook.Close(0);
                while (Marshal.ReleaseComObject(xlWorkBook) > 0) { }
                xlWorkBook = null;
                while (Marshal.ReleaseComObject(xlsWorkSheet) > 0) { }
                xlsWorkSheet = null;
                GC();
                excel.Quit();
                xlApp.Quit();
                while (Marshal.ReleaseComObject(excel) > 0) { }
                excel = null;
                while (Marshal.ReleaseComObject(xlApp) > 0) { }
                excel = null;
                GC();
                return 1;
            }
            catch { return 0; }
        }

        #endregion ExcelModifierTemp

        #region ExcelProcessorSpecial

        public void PreprocessExcelSpcl(string FileAddr)
        {
            ExcelSpecial.Workbook wb;
            ExcelSpecial.Worksheet ws;


            if (!string.Equals(FileAddr, "Nothing Selected"))//NoFileFound InputFilename[counter]
            {
                wb = excel.Workbooks.Open(FileAddr);
                ws = wb.Worksheets[1];//sheet

                int sheetColumn = 1;
                int sheetRow = 1;

                //string trimAddress = InputFilename[counter].Substring(0, InputFilename[counter].Length - 3);
                //public List<string> DatDBColumnNames = new List<string>();
                //public List<string> FileColumnNames = new List<string>();
                //public List<string> ValueForColumnStatics = new List<string>();
                //public List<string> MergeEvents = new List<string>();
                string TempDynamicString_XML;
               for (int count = 12;count<16;count++)
                {
                    try
                    {
                        TempDynamicString_XML = ws.Cells[count, 1].Value2 + string.Empty;
                        if(!TempDynamicString_XML.Contains("Place"))
                            flag_XML_PLACE = false;
                    }
                    catch { flag_XML_PLACE = false; }
                }

                while (ws.Cells[sheetRow + 1, sheetColumn].Value != null)
                {
                    try
                    {
                        dynamic TempDynamicString1;
                        try
                        {
                            TempDynamicString1 = ws.Cells[sheetRow, sheetColumn].Value2;
                            DatDBColumnNames.Add(TempDynamicString1 + string.Empty);
                        }
                        catch { }

                        try
                        {
                            TempDynamicString1 = ws.Cells[sheetRow + 1, sheetColumn].Value2;
                            FileColumnNames.Add(TempDynamicString1 + string.Empty);
                        }
                        catch { }

                        try
                        {
                            TempDynamicString1 = ws.Cells[sheetRow + 2, sheetColumn].Value2;
                            ValueForColumnStatics.Add(TempDynamicString1 + string.Empty);
                        }
                        catch { }

                        try
                        {
                            TempDynamicString1 = ws.Cells[sheetRow + 4, sheetColumn].Value2;
                            MergeEvents.Add(TempDynamicString1 + string.Empty);
                        }
                        catch { }
                        ////////////////
                        if(flag_XML_PLACE)
                        {
                            try
                            {
                                TempDynamicString1 = ws.Cells[sheetRow + 11, sheetColumn].Value2;
                                Place1.Add(TempDynamicString1 + string.Empty);
                            }
                            catch { }

                            try
                            {
                                TempDynamicString1 = ws.Cells[sheetRow + 12, sheetColumn].Value2;
                                Place2.Add(TempDynamicString1 + string.Empty);
                            }
                            catch { }

                            try
                            {
                                TempDynamicString1 = ws.Cells[sheetRow + 13, sheetColumn].Value2;
                                Place3.Add(TempDynamicString1 + string.Empty);
                            }
                            catch { }

                            try
                            {
                                TempDynamicString1 = ws.Cells[sheetRow + 14, sheetColumn].Value2;
                                Place4.Add(TempDynamicString1 + string.Empty);
                            }
                            catch { }
                        }
                        //////////////////////
                        sheetColumn++;
                    }
                    catch
                    {
                        sheetColumn++;
                    }

                }
                flag_XML_PLACE = true;
                wb.Close(0);
            }
        }
        #endregion ExcelProcessorSpecial

        #region WriteCSVSpecial_Tab_6

        public void WriteCSVSpecial(List<string> Spcl_FileColumnNames, int RowCounter, dynamic[,] ArrayMessageFromDatabase, [Optional] string Path)//, List<string> RowValue,string Path
        {
            string tempString = string.Empty;
            for (int row = 0; row <= RowCounter; row++)
            {
                for (int column = 0; column < Spcl_FileColumnNames.Count; column++)
                {
                    try
                    {
                        if (row == 0 && !Spcl_FileColumnNames[column].Equals("."))
                        {

                            if (column == (Spcl_FileColumnNames.Count - 1))
                                File.AppendAllText(Path, Spcl_FileColumnNames[column]);
                            else
                                File.AppendAllText(Path, Spcl_FileColumnNames[column] + ",");
                        }

                        else
                        {
                            if (!string.IsNullOrEmpty(ArrayMessageFromDatabase[row - 1, column]))
                            {
                                if (column == (Spcl_FileColumnNames.Count - 1))
                                    tempString += ArrayMessageFromDatabase[row - 1, column];
                                else
                                    tempString += ArrayMessageFromDatabase[row - 1, column] + ",";
                            }
                            else
                            {
                                if (column == (Spcl_FileColumnNames.Count - 1))
                                    tempString += ArrayMessageFromDatabase[row - 1, column];
                                else
                                    tempString += ",";
                            }
                        }

                    }
                    catch { }
                }
                if(!string.IsNullOrEmpty(tempString))
                    File.AppendAllText(Path, tempString+"\r\n"); tempString = string.Empty;
            }
        }


        #endregion WriteCSVSpecial

        #region WriteXMLSpecial_Tab_6

        public void WriteXMLSpecial(List<string> Spcl_FileColumnNames, int RowCounter, dynamic[,] ArrayMessageFromDatabase, List<string> place1, List<string>place2, List<string> place3, List<string> place4,string DeviceType, [Optional] string Path)
        {

            string tempString = string.Empty;


            using (XmlWriter writer = XmlWriter.Create(Path))
            {
                bool flag_InLoop = false; bool flag_IsHeaderchanged = false; bool flag_EndOfSet_CloseAll = false;

                writer.WriteStartElement("XML"); writer.WriteString("\r\n");

                    writer.WriteElementString("DEVICETYPE", DeviceType); writer.WriteString("\r\n");

                writer.WriteStartElement(place2[0]); writer.WriteString("\r\n");  //Header  
                int count = 0;
                    while (place2[count].Contains(place2[0]))
                    {
                        try
                        {
                            writer.WriteElementString(Spcl_FileColumnNames[count] + string.Empty, ArrayMessageFromDatabase[0, count] + string.Empty);
                            writer.WriteString("\r\n");
                        }
                        catch(Exception c){ MessageBox.Show(string.Empty + c); }
                        count++;
                    }
                writer.WriteEndElement(); //end Header writer.WriteString("/"+place2[0]);
                writer.WriteString("\r\n");
                writer.WriteStartElement("DEVICES"); writer.WriteString("\r\n");  //Devices
                count = 0;



                for (int row = 0; row <= RowCounter; row++)
                {
                    int KWHCounter = 0; flag_EndOfSet_CloseAll = false;

                    writer.WriteStartElement("DEVICE"); writer.WriteString("\r\n");

                    for (int column = 0; column < Spcl_FileColumnNames.Count; column++)
                    {
                        
                        if(!string.IsNullOrEmpty(place3[column]))       //if not empty
                        {
                            if (!string.IsNullOrEmpty(place4[column]))      //if not empty
                            {
                                if(KWHCounter==0 && place4[column].Contains("KWHTEST"))
                                {
                                    writer.WriteStartElement("KWHTEST"); writer.WriteString("\r\n");
                                    flag_InLoop = true; flag_IsHeaderchanged = true;
                                     KWHCounter++;
                                }
                                if (place4[column].Contains("COMMUNICATIONDEVICES"))
                                {
                                    if(flag_IsHeaderchanged)
                                    {
                                        writer.WriteEndElement(); writer.WriteString("\r\n"); flag_InLoop = false;      //kwh test
                                        flag_IsHeaderchanged = false;
                                        writer.WriteStartElement(place4[column]); writer.WriteString("\r\n");   //comm
                                        writer.WriteStartElement("DEVICE"); writer.WriteString("\r\n");
                                        flag_InLoop = true;
                                        KWHCounter++;
                                    }
                                }
                                if (place4[column].Contains("END"))
                                {
                                    flag_EndOfSet_CloseAll = true;
                                    writer.WriteEndElement(); writer.WriteString("\r\n");
                                    writer.WriteEndElement(); writer.WriteString("\r\n");
                                    writer.WriteEndElement(); writer.WriteString("\r\n");
                                    //writer.Flush();
                                }
                                if(!flag_EndOfSet_CloseAll)
                                {
                                    writer.WriteElementString(Spcl_FileColumnNames[column] + string.Empty, ArrayMessageFromDatabase[row, column] + string.Empty);
                                    writer.WriteString("\r\n");
                                }
                            }
                            else
                            {
                                writer.WriteElementString(Spcl_FileColumnNames[column] + string.Empty, ArrayMessageFromDatabase[row, column] + string.Empty); writer.WriteString("\r\n");
                            }
                        }
                        count++; 
                    }
                }
                writer.WriteEndElement(); writer.WriteString("\r\n");
                writer.Flush();
                writer.Close();
            }

        }


        #endregion WriteXMLSpecial


        #region EXCELoptionForSpecial
        public void WriteExcelSpecial(List<string> Spcl_FileColumnNames, int RowCounter, dynamic[,] ArrayMessageFromDatabase, [Optional] string Path)//, List<string> RowValue,string Path
        {
            //We have to creatre an option to choose the format.

            _Excel.Workbook xlWorkBookS;
            _Excel.Worksheet xlWorkSheetS;
            object misValue = System.Reflection.Missing.Value;


            //xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkBookS = excel.Workbooks.Add(misValue);
            xlWorkSheetS = (_Excel.Worksheet)xlWorkBookS.Worksheets.get_Item(1);
            //XMLParser x = new XMLParser();

            for (int column = 0; column < Spcl_FileColumnNames.Count; column++)
            {
                xlWorkSheetS.Columns[column + 1].NumberFormat = "@";
            }

            for (int row = 0; row <= RowCounter; row++)
            {
                for (int column = 0; column < Spcl_FileColumnNames.Count; column++)
                {
                    try
                    {
                        if (row == 0)
                        {
                            xlWorkSheetS.Cells[row + 1, column + 1] = Spcl_FileColumnNames[column];
                            xlWorkSheetS.Cells[row + 1, column + 1].Font.Bold = true;
                            xlWorkSheetS.Rows[1].Interior.Color = System.Drawing.Color.Gray;
                        }

                        else
                        {
                            if (!Spcl_FileColumnNames[column].ToUpper().Contains("IMEI") && !Spcl_FileColumnNames[column].ToUpper().Contains("SIMCARDID") && ArrayMessageFromDatabase[row - 1, column].Length >= 13)
                                xlWorkSheetS.Cells[row + 1, column + 1] = "\'" + ArrayMessageFromDatabase[row - 1, column];
                            else
                                xlWorkSheetS.Cells[row + 1, column + 1] = ArrayMessageFromDatabase[row - 1, column];
                        }
                    }
                    catch { }
                }
            }

            try
            {
                xlWorkBookS.SaveAs(Path);
            }
            catch { MessageBox.Show("The Excel File is not being Saved!", "Warning>> \r\n", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            xlWorkBookS.Close(0);

            while (Marshal.ReleaseComObject(xlWorkBookS) > 0) { }
            xlWorkBookS = null;
            while (Marshal.ReleaseComObject(xlWorkSheetS) > 0) { }
            xlWorkSheetS = null;
            GC();
            excel.Quit();
            xlApp.Quit();
            while (Marshal.ReleaseComObject(excel) > 0) { }
            excel = null;
            while (Marshal.ReleaseComObject(xlApp) > 0) { }
            excel = null;
            GC();
        }
        #endregion ExcelOptionForSpecial
        #region Comment
        //all the functions are here just expand this to see more.

        //if(row ==0 && ColumnValue[column].Contains("IMEI"))
        //    xlWorkSheet.Columns[column+1].NumberFormat = "#########################";//356 441 114 107 465

        //if (row == 0 && ColumnValue[column].Contains("SimCardID"))
        //{
        //    xlWorkSheet.Columns[column+1].NumberFormat = "@";//Pasting it as a string. other way around is to add ' in front of every number to keep it as it is or to avoid rounding off.
        //    //xlWorkSheet.Columns[column].NumberFormat = "@";//"#########################";//89 148 000 005 518 800 000
        //}
        //if (row == 0 && ColumnValue[column].Contains("KwhUsage"))
        //{
        //    xlWorkSheet.Columns[column + 1].NumberFormat = "@";//Pasting it as a string. other way around is to add ' in front of every number to keep it as it is or to avoid rounding off.
        //}


        //#region ReadCell
        //public void ReadCell(int sheetRow)
        //{
        //    int sheetColumn = 1;
        //    string fileName = @"c:\Temp\ExcelLayout.xml";
        //    while (ws.Cells[sheetRow,sheetColumn].Value !=null)
        //    {
        //        Dataset.Add(ws.Cells[sheetRow, sheetColumn].Value2);
        //        if (ws.Cells[sheetRow + 1, sheetColumn].Value2 != null)
        //        {
        //            dynamic demo = ws.Cells[2, sheetColumn].Value2;
        //            DatasetRow2.Add(demo+string.Empty);
        //        }
        //        else
        //            DatasetRow2.Add(string.Empty);
        //        sheetColumn++;
        //    }
        //    ExcelLayout layout = ExcelLayoutManager.Initialize(Dataset,DatasetRow2);
        //    ExcelLayoutManager.Save(fileName, layout, true);
        //}
        //#endregion ReadCell

        //#region AddCells
        //public void AddCells(List<int> temp1, List<string> temp2,string File1Name)
        //{
        //    #region CommentedCode
        //    //ColumnNumberToAddFromFile2
        //    //ColumnNameToAddFromFile2
        //    //_Application excel = new _Excel.Application();
        //    //Workbook wb;
        //    //Worksheet ws;
        //    #endregion CommentedCode

        //    int counter = 0;
        //    try
        //    {
        //        while (counter<temp1.Count)
        //        {
        //            ws.Columns[temp1[counter]].Cells.Insert();//ws.Columns[temp1[counter]+counter].Cells.Insert();
        //            ws.Cells[1,temp1[counter]].Value2 = temp2[counter]+"."+counter;
        //            #region CommentedCode
        //            //this works as well// this does not insert the column into the sheet it just replaces the name it has previously.//ws.Cells[1,temp1[counter] + counter].Value2 = "Cell"+ temp1[counter]+"."+counter; 
        //            //ws.Columns.
        //            //ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
        //            ////wb.Save();
        //            ////wb.Close();
        //            //excel.Workbooks.Close();
        //            #endregion CommentedCode

        //            counter++;
        //        }
        //        ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
        //        wb.Close();
        //        #region CommentedCode
        //        //excel.Workbooks.Close();
        //        //ws.Columns["D"].Clear();//Works
        //        //ws.Columns["D"].Delete();//works

        //        //ws.Columns[2].Cells.Insert();
        //        //ws.Cells[2,2].Value2 = "Cell B2."; //this works as well// this does not insert the column into the sheet it just replaces the name it has previously.
        //        //ws.Rows.Cells[1,1].Value2 = "column B.";//works-> row column
        //        ////ws.Columns.Insert(3);
        //        ////ws.Columns[demo].Insert();
        //        //ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
        //        //wb.Save();
        //        //wb.Close();
        //        //excel.Workbooks.Close();
        //        #endregion CommentedCode
        //    }
        //    catch (Exception ex){ MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error); }

        //}
        //#endregion AddCells

        //#region DeleteCells
        //public void DeleteCells(List<int> temp1, List<string>temp2, string File1Name)
        //{
        //    int counter = 0;
        //    while (counter<temp1.Count)
        //    {
        //        ws.Columns[temp1[counter]+1].Clear();//Works
        //        //ws.Columns[temp1[counter]+1].Delete();//works
        //        counter++;
        //    }
        //    ws.SaveAs(@"G:\_ShipmentFiles\vishalModified\" + File1Name);
        //    #region CommentedCode
        //    //excel.Workbooks.Close();
        //    //int counter = temp1.Count-1;
        //    //while (counter > 0)
        //    //{
        //    //    ws.Columns[temp1[counter]].Add();
        //    //}
        //    //ws.Cells.EntireColumn.Delete(1);
        //    //ws.Columns[3].Clear();
        //    //ws.Columns[3].Delete();
        //    //wb.Save(); wb.Close();
        //    #endregion CommentedCode
        //}
        //#endregion DeleteCells

        //#region EmptyCellsCheck
        //public void CheckForEmptyCells(List<int> CNumberDFF1, List<string> CNameDFF1)
        //{
        //    //ColumnNumberToDeleteFromFile1     CNumberDFF1
        //    //ColumnNameToDeleteFromFile1       CNameDFF1


        //    for (int counter=0;counter< CNumberDFF1.Count;)
        //    {
        //        //can be implemented
        //        if (ws.Cells[2, CNumberDFF1[counter]].Value2 != null)
        //        {
        //            CNumberDFF1.RemoveAt(counter);
        //            CNameDFF1.RemoveAt(counter);
        //            if (counter != 0) { counter--; }
        //            else { counter = 0; }
        //        }
        //        counter++;
        //    }
        //}
        //#endregion EmptyCellsCheck-+

        //#region CleanNamelessColumns
        //public void CleanNamelessColumns(List<int> CNumberDFF1)
        //{
        //    int FlagtoExit = 0;
        //    for (int counter = 1; counter < 200;)
        //    {
        //        //can be implemented
        //        if (ws.Cells[1, counter].Value2 == null)
        //        {
        //            ws.Columns[counter].Delete();
        //            if (counter != 0) { counter--; FlagtoExit++; }
        //            else { counter = 0; }
        //        }
        //        else { FlagtoExit=0; }
        //        counter++;
        //        if (FlagtoExit >= 10)//FlagtoExit >= 10//FlagtoExit >= CNumberDFF1.Count
        //            break;
        //    }
        //    wb.Save(); excel.Workbooks.Close();
        //}
        //#endregion CleanNamelessColumns

        //#region ExcelExtraction
        //public void ExcelExtraction(List<string>InputFilename,List<string>FileDir,List<string>DirectoryName)
        //{

        //    for (int counter =0;counter< FileDir.Count; counter++)// FileDir.Count
        //    {
        //        wb = excel.Workbooks.Open(FileDir[counter]);
        //        ws = wb.Worksheets[1];//sheet

        //        int sheetColumn = 1;
        //        int sheetRow = 1;
        //        string trimAddress = InputFilename[counter].Substring(0,InputFilename[counter].Length-3);

        //        Directory.CreateDirectory(@"F:\ShipmentsXMLfiles\" + DirectoryName[counter]);
        //        string OutputFileName = @"F:\ShipmentsXMLfiles\" + DirectoryName[counter] + "\\" + trimAddress+"xml"; //chjange the location
        //        while (ws.Cells[sheetRow, sheetColumn].Value != null)
        //        {
        //            try
        //            {
        //                dynamic TempDynamicString1 = ws.Cells[sheetRow, sheetColumn].Value2;
        //                Dataset.Add(TempDynamicString1 + string.Empty);
        //                if (ws.Cells[sheetRow + 1, sheetColumn].Value2 != null)
        //                {
        //                    dynamic TempDynamicString2 = ws.Cells[2, sheetColumn].Value2;
        //                    DatasetRow2.Add(TempDynamicString2 + string.Empty);
        //                }
        //                else
        //                    DatasetRow2.Add(string.Empty);
        //                sheetColumn++;
        //            }
        //            catch
        //            {
        //                sheetColumn++;
        //            }

        //        }
        //        ExcelLayout layout = ExcelLayoutManager.Initialize(Dataset, DatasetRow2);
        //        ExcelLayoutManager.Save(OutputFileName, layout, true);
        //        wb.Close();
        //        Dataset.Clear();
        //        DatasetRow2.Clear();
        //    }
        //}
        //#endregion XLSExtraction

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
        #endregion Comment

    }
}

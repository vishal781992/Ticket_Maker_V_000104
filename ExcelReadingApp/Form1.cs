using System;
using System.Collections.Generic;
//using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.Linq;
//using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.ComponentModel;
//using Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel;
//using System.Windows.Forms;
//using Microsoft.VisualBasic.FileIO;
using System.IO;
using System.Diagnostics;
using System.Dynamic;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.CodeDom;
using System.ComponentModel.Design;
//using System.Diagnostics;

namespace ExcelReadingApp
{
    public partial class Form1 : Form  
    {
        #region Declarations
        DeclarationClass DC = new DeclarationClass();

        public string FileNameFromRootDir { get; set; }
        //public string FileInputDir = @""; //the file location changer meanwhile debugging
       

        string[] users = new string[] { "vishal", "steve" };
        string[] Dte = new string[3500];
        string[] AryOfColumns = new string[3500];

        public dynamic[,] Spcl_ArrayMessageFromDatabase = new dynamic[4000, 200];


        public const int APPEND = 1,
                         NEWLine = 2;

        public int counterForFileGeneratedInXml;



        public List<int> ColumnNumberToAddFromFile2 = new List<int>(),
                         ColumnNumberToDeleteFromFile1 = new List<int>();

        public List<string> ColumnNameToDeleteFromFile1 = new List<string>(),
                            DatabaseList = new List<string>(),
                            TempList = new List<string>(),
                            TempList2 = new List<string>(),
                            TicketsListForDataQuerySQL = new List<string>(),
                            TicketNumberIndividual = new List<string>(),
                            ColumnNameToAddFromFile2 = new List<string>(),
                            SearchDataTab4 = new List<string>(),
                            LogDirNames = new List<string>(),

                            Spcl_DatDBColumnNames = new List<string>(),
                            Spcl_FileColumnNames = new List<string>(),
                            Spcl_ValueForColumnStatics = new List<string>(),
                            Spcl_MergeEvents = new List<string>(),

                            Place1 = new List<string>(),
                            Place2 = new List<string>(),
                            Place3 = new List<string>(),
                            Place4 = new List<string>();

        DateTime StartDate,
                 EndDate;

        RootDirectoriesExplorer RE1 = new RootDirectoriesExplorer();
        MessageBox_User MBU = new MessageBox_User();

        #endregion Declarations

        #region Form init
        public Form1()
        {
            InitializeComponent();
            myBackgroundWorker = new BackgroundWorker();
            myBackgroundWorker.WorkerReportsProgress = true;
            myBackgroundWorker.WorkerSupportsCancellation = false;
            myBackgroundWorker.DoWork += myBackgroundWorker1_DoWork;
            myBackgroundWorker.RunWorkerCompleted += myBackgroundWorker1_RunWorkerCompleted;
            myBackgroundWorker.ProgressChanged += myBackgroundWorker1_ProgressChanged;

            myBackgroundWorkerTab3 = new BackgroundWorker();
            myBackgroundWorkerTab3.WorkerReportsProgress = true;
            myBackgroundWorkerTab3.WorkerSupportsCancellation = false;
            myBackgroundWorkerTab3.DoWork += myBackgroundWorkerTab3_DoWork;
            myBackgroundWorkerTab3.RunWorkerCompleted += myBackgroundWorkerTab3_RunWorkerCompleted;
            myBackgroundWorkerTab3.ProgressChanged += myBackgroundWorkerTab3_ProgressChanged;

            myBackgroundWorkerTab4 = new BackgroundWorker();
            myBackgroundWorkerTab4.WorkerReportsProgress = true;
            myBackgroundWorkerTab4.WorkerSupportsCancellation = false;
            myBackgroundWorkerTab4.DoWork += myBackgroundWorkerTab4_DoWork;
            myBackgroundWorkerTab4.RunWorkerCompleted += myBackgroundWorkerTab4_RunWorkerCompleted;
            myBackgroundWorkerTab4.ProgressChanged += myBackgroundWorkerTab4_ProgressChanged;

            myBackgroundWorkertab6 = new BackgroundWorker();
            myBackgroundWorkertab6.WorkerReportsProgress = true;
            myBackgroundWorkertab6.WorkerSupportsCancellation = false;
            myBackgroundWorkertab6.DoWork += myBackgroundWorkertab6_DoWork;
            myBackgroundWorkertab6.RunWorkerCompleted += myBackgroundWorkertab6_RunWorkerCompleted;
            myBackgroundWorkertab6.ProgressChanged += myBackgroundWorkertab6_ProgressChanged;


            this.Text = "Vision TickerMaker " + DC.Version;
        }
        #endregion Form init

        #region Form_Loading
        private void Form1_Load(object sender, EventArgs e)
        {
            /*Note: This commented code below is important if you need to load the config file. Write now it seems to be overkilling and hence does not need. 
             */
            //string StartUppath = System.IO.Directory.GetCurrentDirectory(); StartUppath = Directory.GetParent(StartUppath).Parent.Parent.FullName;
            RTB1.AppendText("The Process is Simple," +
                                        "\r\n1. Enter the Ticketnumber from the Ticket you are holding(Top right corner)" +
                                        "\r\n3. Enter the Database keywoard if it is not a common database listed below." +
                                        "\r\nSelect from the same drop down menu." +
                                        "\r\n4. Enter the PO# number." +
                                        "\r\n6.Hit Start." +
                                        "\r\nThank you!" +
                                        "\r\nVersion " + DC.Version + " includes\r\n" + DC.VersionDetails);

            checkBox_KeepTheLog.Checked = true;
            checkBox_SupressWarnings.Checked = false;

            label_TicketNumberDisplay.Visible = false;
            label_Date.Text = "Date: " + GetCurrentDateAndTime(false);
            labelVersionNumber1.Text = DC.Version;
            monthCalendarStart.Visible = false; monthCalendarEnd.Visible = false; label_Database.Visible = false; button_Refresh.Visible = false; //button invisible for a while
            myBackgroundWorker.RunWorkerAsync(2); this.PBar1.Maximum = 500; //Button_main.Visible = false;
            toolTip_version.SetToolTip(labelVersionNumber1, DC.VersionDetails);
            checkBox_tab1_Intellicode.Checked = true;
            //textBox_PickTicketNumber.Text = "195006,194999,195018,195013,195015,195011,195016,195010,195019,194997,194998"; //only fro debug, else clean it.
            labeltab5_1.Visible = false; label32.Visible = false;
            checkBox_tab1_Intellicode.Visible = false; checkBox_KeepTheLog.Visible = false; checkBox_SupressWarnings.Visible = false; checkBox_SupressWarnings.Checked = true;
            richTextBox_5.Visible = false;
            textBoxT5_SearchTB.Visible = false;
            buttonT5_Search.Visible = false;


            //only fro debug, else clean it.
            button_ForDebug.Visible = false;
    }
        #endregion Form_Loading

        #region Tab 1 Start Button 
        private void ButtonClick_Start(object sender, EventArgs e)
        {
            label_TicketNumberDisplay.Visible = true; TicketNumberIndividual.Clear(); TicketsListForDataQuerySQL.Clear();
            DC.Log_TicketToLog = string.Empty; DC.Log_TicketCounts = string.Empty;
            PBar1.Value = 0;PBar1.Maximum = 500;
            DC.flag_POnumber = true; label_TicketNumberDisplay.Text = textBox_PickTicketNumber.Text;

            #region intellicode checked?
            if (checkBox_tab1_Intellicode.Checked)//when the automatic column detection is enabled
            {
                richTextBox_FileFormat.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage," +
                       "AlternateID,PreviousID,IMEI,SimCardID,\r\nDevEUI,CommID,8digitCommID,CommID1,CommID2,CommID3,CommID4,ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
            }
            #endregion intellicode checked?

            #region SalesPerson
            if (string.Equals(comboBox_SalesPerson.Text, "SalesPerson") || string.IsNullOrEmpty(comboBox_SalesPerson.Text))  //combobox Salesperson
                comboBox_SalesPerson.ForeColor = Color.Red;
            #endregion SalesPerson

            if (!string.IsNullOrEmpty(textBox_PickTicketNumber.Text) && !string.IsNullOrEmpty(textBox_CustomerPO.Text) && !string.Equals(FileNameFromRootDir, "ERROR SELECTION")
                  && !string.IsNullOrEmpty(comboBox_CompanyName.Text) && !string.IsNullOrEmpty(comboBox_DataBaseName.Text))
            {
                #region inititial declaration
                PBar1.PerformStep();

                RTB1.AppendText("\r\nThe START button is pressed, wait for the program to create a File for you." +
                    "\r\nA Message will popup as the File is created successfully(If not Supressed). If you  get a message(Popup) for Replacing the existing file" +
                    "\r\nHit YES if you want to overwrite, Else NO!");

                ExcelProcessor EX = new ExcelProcessor();
                FormatModifier FM = new FormatModifier();
                QueryTest QT = new QueryTest();

                PBar1.PerformStep();
                FM.FormatString = richTextBox_FileFormat.Text;

                #endregion inititial declaration

                #region Program Sequence
                /*format Parser function performs sorting the manual entry of the format entered. This function removes "," and places them into a List Column Values which is used in the program.
                 * if this function is performed it sets the flag(Flag_UseXMLLoadDataFun) so that xmlLoadData will skip.
                 */
                FM.FormatParser();

                PBar1.Value += 10;//1

                FM.xmlLoadData(FileNameFromRootDir);

                if (FM.ColumnValue.Count > 0)
                {
                    PBar1.Value += 10;

                    /*Note: XMLRequestData, this function load the original file to take the company name and not takes from the user, remember it uses the user defined format but
                     * for company name it loads the xml file in the records given through the "Sold to" text input
                     */
                    string CompanyName = FM.XMLRequestData("Company", DC.FilePathOfXML);
                    DC.databaseType = comboBox_DataBaseName.Text.ToUpper().EndsWith("VISION") ? "dbo" : "power";

                    /*USER_init function is used to init all the necessary credentials for the test query SQL. This is not to be messed up with.
                     * Everything about the SQl should be happen before this is executed.
                     */
                    QT.USER_init(comboBox_DataBaseName.Text);
                    PBar1.Value += 10;

                    /*Ticket_Formater helpful in multiple ticket entries. it removes any delimiter in the string and adds it to the list.*/

                    Ticket_Formater(textBox_PickTicketNumber.Text, DC.databaseType);

                    QT.Tab1_TestQuery(FM.ColumnValue, CompanyName, TicketsListForDataQuerySQL, textBox_CustomerPO.Text, DC.databaseType);//database type is important here

                    PBar1.Value += 10;//4

                    string CompletePathForXLSXexport = string.Empty; string FileNameExtension = string.Empty;

                    //saving into _Shipments
                    if (checkBox_Tab1_SaveInShipment.Checked)    
                    {
                        if (TicketNumberIndividual.Count > 1)
                        {
                            //Multiple tickets are processed here
                            FileNameExtension = CompanyName + "_PT" + TicketNumberIndividual[0] + "_M" + TicketsListForDataQuerySQL.Count + "_PO" +
                            textBox_CustomerPO.Text + "_" + GetCurrentDateAndTime(true) + ".xlsx";

                            //for log purpose
                            DC.Log_TicketCounts = "M" + TicketsListForDataQuerySQL.Count;

                            CompletePathForXLSXexport = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39) + FileNameExtension;
                            DC.folderNameForOutputFile = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39);//39 is the size of the string

                            
                        }
                        else
                        {
                            //Single ticket
                            FileNameExtension = CompanyName + "_PT" + TicketNumberIndividual[0] + "_PO" +
                            textBox_CustomerPO.Text + "_" + GetCurrentDateAndTime(true) + ".xlsx";

                            //Log purpose
                            DC.Log_TicketCounts = "S";

                            CompletePathForXLSXexport = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39) + FileNameExtension;
                            DC.folderNameForOutputFile = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39);
                        }
                    }
                    //saving into Vishal_Shipments(demo folder for testing)
                    else
                    {
                        if (TicketNumberIndividual.Count > 1)
                        {
                            FileNameExtension = CompanyName + "_PT" + TicketNumberIndividual[0] + "_M" + TicketsListForDataQuerySQL.Count + "_PO" +
                            textBox_CustomerPO.Text + "_" + GetCurrentDateAndTime(true) + ".xlsx"; DC.Log_TicketCounts = "M" + TicketsListForDataQuerySQL.Count;
                            CompletePathForXLSXexport = DC.ExportXlSXPath + FileNameExtension;
                            DC.folderNameForOutputFile = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39);
                        }
                        else
                        {
                            FileNameExtension = CompanyName + "_PT" + TicketNumberIndividual[0] + "_PO" +
                           textBox_CustomerPO.Text + "_" + GetCurrentDateAndTime(true) + ".xlsx"; DC.Log_TicketCounts = "S";
                            CompletePathForXLSXexport = DC.ExportXlSXPath + FileNameExtension;
                            DC.folderNameForOutputFile = DC.ORIGINALSHIPMENTPATH_ + DC.ExportXlSXPath.Substring(39);
                        }
                    }
                    PBar1.Value += 10;//5
                    //dynamic[,] demo = QT.ArrayMessageFromDatabase;

                    /*WriteANewExcel, this function writes the new Excel File to the Directory. This handles writing all the rows and columns to the file.
                     * CompletePathForXLSXexport responsible for file path and extension.
                     */
                    EX.WriteANewExcel(FM.ColumnValue, QT.RowCounter, QT.ArrayMessageFromDatabase, CompletePathForXLSXexport);//CompanyName_TicketNumber_DataBase_Date

                    /*Removes unwanted columns from the existing Excel file created above. Important function
                     */
                    if(!checkBox_tab1_deleteEmptyCinXLS.Checked)
                    {
                        #region Excel Modification
                        int result = EX.ExcelModifierFunction(CompletePathForXLSXexport, FM.ColumnValue.Count);
                        if(result==1)
                            RTB1.AppendText("Excel Intelligent column detection is done.");
                        else
                            RTB1.AppendText("Error in column detection process.");
                        #endregion Excel Modification
                    }
                    PBar1.Value += 10;//6 
                    DC.FileNameExtension_Global = FileNameExtension;
                    RTB1.Text = "File name: " + FileNameExtension+ "\r\n";  //start of the richtext box text
                    RTB1.AppendText("\r\nFolder name: " + DC.folderNameForOutputFile + "\r\n\r\n");
                    RTB1.AppendText( QT.RowCounter + " VM --> " + CompanyName + "," + " PO#:" + textBox_CustomerPO.Text +
                        "," + "PT:" + textBox_PickTicketNumber.Text+", "+"DB: "+ comboBox_DataBaseName.Text+".");

                    RTB1.AppendText(DeclarationClass.DividerString);
                #endregion Program Sequence

                #region DataVerificationForUser

                    dataVerification DV = new dataVerification(QT.ArrayMessageFromDatabase, FM.ColumnValue, QT.RowCounter, QT.MeterTypeCodes);
                    try
                    {
                        foreach (string TicketNumber in TicketNumberIndividual)
                        {/*This Verification_ItemRange and Verification_General_typeSort functions verify the data according to the tickets standards to the user by displaying them to the screen.
                      * Next task is to keep the log of the files. according to the 
                      */
                            RTB1.AppendText("\r\nMeter Range: " + DV.Verification_ItemRange("MeterID", TicketNumber));

                            #region Pallet, Box
                            RTB1.SelectionColor = Color.Red;
                            RTB1.AppendText("\r\n(B-)Box: " + DV.Verification_ItemRange("Box", TicketNumber));
                            RTB1.AppendText("\r\n(P-)Pallet: " + DV.Verification_ItemRange("Pallet", TicketNumber)+ "  |  Blank Box/s: "+DV.Flag_ErrorInPallet);
                            RTB1.SelectionColor = Color.Black;
                            #endregion Pallet, Box

                            //this is general sorting methods used, supply the name and it gives us the output
                            #region FirmwareRevision
                            RTB1.AppendText("\r\nFW: ");
                            TempList = DV.Verification_General_typeSort("FirmwareRevision", TicketNumber);
                            foreach (string stringin in TempList)
                                RTB1.AppendText(stringin + "- "); TempList.Clear();
                            PBar1.Value += 10;//7

                            TempList = DV.Verification_General_typeSort("MeterTypeCode", TicketNumber);
                            #endregion FirmwareRevision

                            #region Meter Classification
                            RTB1.AppendText("\r\nMeter Classification:\r\n"); TempList2.Clear();
                            TempList2 = DV.MeterTypeClassification(TempList); int CounterTemp = 0;
                            foreach (string stringin in TempList)
                            {
                                RTB1.AppendText(stringin + "-("+TempList2[CounterTemp]+")" + "\r\n"); CounterTemp++;
                            }
                            TempList.Clear();
                            #endregion Meter Classification

                            #region Form/base
                            RTB1.AppendText("\r\nForm: ");
                            TempList = DV.Verification_General_typeSort("Form/Base", TicketNumber);
                            foreach (string stringin in TempList)
                                RTB1.AppendText(stringin + "- "); TempList.Clear();
                            #endregion Form/base

                            #region ClassAmps
                            RTB1.AppendText("\r\nClass: ");
                            TempList = DV.Verification_General_typeSort("ClassAmps", TicketNumber);
                            foreach (string stringin in TempList)
                                RTB1.AppendText(stringin + "- "); TempList.Clear();
                            #endregion ClassAmps

                            #region CompanyName
                            RTB1.AppendText("\r\n" + DV.RowCounterForTheSpecTicket + " of " + QT.RowCounter + " VM --> " + CompanyName + "," + " PO#:" + textBox_CustomerPO.Text +
                                                        "," + "PT:" + TicketNumber+"\r\n");
                            TempList.Clear();
                            #endregion CompanyName

                            DV.CounterGenerator();
                            PBar1.Value += 10;//8
                            if (TicketNumberIndividual.Count > 1)
                                RTB1.AppendText(DeclarationClass.DividerString1);
                        }
                        //crossVerify the columns
                        //this step is the last and more time consuming than everyone.

                        #region Verification
                        RTB1.AppendText("\r\n-----------------Verification-------------------- For All Tickets(If Multiple)");

                        #region CommIDs
                        TempList = DV.VerificationOfCommID();                 //commIDS
                        if (TempList.Count > 0)
                        {
                            RTB1.AppendText("\r\nCommID Errors(Dont consider them as Errors as it checks for format matching 05 or 08. Many more formats exists.)");
                            foreach (string stringin in TempList)
                                RTB1.AppendText(stringin + "- "); TempList.Clear();
                        }
                        else
                            RTB1.AppendText("\r\nCommID's, No error");
                        #endregion CommIDs

                        #region ALxx
                        var tuple = DV.Verification_AL_Checks("ALSF");              //ALSF
                        if(tuple.Item1.Count>0)
                        {
                            RTB1.AppendText("\r\nALSF error: ");
                            foreach (string stringin in tuple.Item1)
                                RTB1.AppendText(stringin + "-");
                        }
                        else
                            RTB1.AppendText("\r\nALSF, No error");
                        RTB1.AppendText(", ALSF range: " + tuple.Item2[0] + "<-->" + tuple.Item2[tuple.Item2.Count - 1]); tuple.Item1.Clear(); tuple.Item2.Clear();

                          tuple = DV.Verification_AL_Checks("ALSL");                  //ALSL
                        if (tuple.Item1.Count > 0)
                        {
                            RTB1.AppendText("\r\nALSL error: ");
                            foreach (string stringin in tuple.Item1)
                                RTB1.AppendText(stringin + "-");
                        }
                        else
                            RTB1.AppendText("\r\nALSL, No error");
                        RTB1.AppendText(", ALSL range: " + tuple.Item2[0] + "<-->" + tuple.Item2[tuple.Item2.Count - 1]); tuple.Item1.Clear(); tuple.Item2.Clear();

                        tuple = DV.Verification_AL_Checks("ALSP");                  //ALSP
                        if (tuple.Item1.Count > 0)
                        {
                            RTB1.AppendText("\r\nALSP error: ");
                            foreach (string stringin in tuple.Item1)
                                RTB1.AppendText(stringin + "-");
                        }
                        else
                            RTB1.AppendText("\r\nALSP, No error");
                        RTB1.AppendText(", ALSP range: " + tuple.Item2[0] + "<-->" + tuple.Item2[tuple.Item2.Count - 1]); tuple.Item1.Clear(); tuple.Item2.Clear();


                        tuple = DV.Verification_AL_Checks("ALWA");                  //ALWA
                        if (tuple.Item1.Count > 0)
                        {
                            RTB1.AppendText("\r\nALWA error: ");
                            foreach (string stringin in tuple.Item1)
                                RTB1.AppendText(stringin + "-");
                        }
                        else
                            RTB1.AppendText("\r\nALWA, No error");

                        RTB1.AppendText(", ALWA range: " + tuple.Item2[0] + "<-->" + tuple.Item2[tuple.Item2.Count - 1]); tuple.Item1.Clear(); tuple.Item2.Clear();
                        #endregion ALxx

                        #region Duplicate check
                        RTB1.AppendText("\r\n-----------------Dupli Checks-------------------- For All Tickets(If Multiple)\r\nCommID:" + DV.DuplicateRecordVerification("CommID",comboBox_DataBaseName.Text, DC.databaseType));
                        RTB1.AppendText("\r\nMeterID:" + DV.DuplicateRecordVerification("MeterID", comboBox_DataBaseName.Text, DC.databaseType));
                        try{RTB1.AppendText("\r\nDevEUI:" + DV.DuplicateRecordVerification("DevEUI", comboBox_DataBaseName.Text, DC.databaseType));}catch {}
                        try { RTB1.AppendText("\r\nIMEI:" + DV.DuplicateRecordVerification("IMEI", comboBox_DataBaseName.Text, DC.databaseType)); } catch {}
                        try { RTB1.AppendText("\r\nSimCardID:" + DV.DuplicateRecordVerification("SimCardID", comboBox_DataBaseName.Text, DC.databaseType)); } catch {}

                        try { RTB1.AppendText("\r\nsame DevEUI,Diff Meters:" + DV.DuplicateRecordVerification("MeterID", comboBox_DataBaseName.Text, DC.databaseType , "DevEUI")); } catch {}
                        try { RTB1.AppendText("\r\nsame IMEI,Diff Meters:" + DV.DuplicateRecordVerification("MeterID", comboBox_DataBaseName.Text, DC.databaseType , "IMEI")); } catch {}
                        try { RTB1.AppendText("\r\nsame SimID,Diff Meters:" + DV.DuplicateRecordVerification("MeterID", comboBox_DataBaseName.Text, DC.databaseType , "SimCardID")); } catch {}
                        RTB1.AppendText("\r\n-------------------------------------------------");
                        #endregion Duplicate check

                        #region keep the log checkbox
                        if (checkBox_KeepTheLog.Checked)
                        {
                            DataLogging DLT = new DataLogging();

                            DC.Log_DataCollectionString = RTB1.Text;
                            foreach (string TicketNum in TicketNumberIndividual)
                                DC.Log_TicketToLog += "<" + TicketNum + "> ";
                            try { DLT.FileOpener(DC.Log_TicketToLog, DC.Log_TicketCounts, DC.Log_DataCollectionString); }
                            catch { RTB1.AppendText("\r\nError in the data Logging"); if (!checkBox_SupressWarnings.Checked) { MessageBox.Show("Important-- Log is not recorded due to some Error."); } }
                        }
                        #endregion keep the log checkbox

                        PBar1.Value += 10;
                        #endregion Verification

                        #region SalesPerson
                        MBU.MB_TextDisplay(DeclarationClass.NotificationString1);
                        if (!string.Equals(comboBox_SalesPerson.Text, "SalesPerson"))
                        {
                            MBU.MB_TextAppend("\r\nEmail to: tom@visionmetering.com" + "\r\nCC: " + SalespersonQuiz(comboBox_SalesPerson.Text));

                        }

                        #endregion SalesPerson

                        #region datalogging Notification
                        MBU.MB_TextAppend("\r\nSuccess, data Logging.");
                        DialogResult dialogR = MBU.ShowDialog();

                        if (dialogR == DialogResult.OK || dialogR == DialogResult.Cancel){}

                        PBar1.Value = PBar1.Maximum;
                        #endregion datalogging Notification
                    }
                    catch (Exception ex)
                    {
                        RTB1.AppendText(DeclarationClass.NotificationString2+Environment.NewLine + ex);
                    }
                    richTextBox_FileFormat.Text = string.Empty;
                }
                #endregion DataVerificationForUser
            }
            else
            {
                RTB1.Text = DeclarationClass.ErrorString1;
                PBar1.Value = PBar1.Maximum;
            }
            Clipboard.SetText(RTB1.Text);
        }
       
        #endregion Tab 1 Start Button

        #region Adjoining Functions

        #region Tab 1
        private void button_CompanyNameConfirm_Click(object sender, EventArgs e)
        {
            RE1.CompanyFinder(comboBox_CompanyName.Text);
            comboBox_CompanyName.DataSource = RE1.FilenamesForSearch;
            comboBox_CompanyName.BackColor = Color.Red;
        }
        private void comboBox_CompanyName_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(comboBox_CompanyName.Text))
                {
                    comboBox_CompanyName.BackColor = Color.LightGreen;
                    DC.IamPopUP = true;
                    DC.Flag_forDisplayOfDatabase = true;
                }
                else
                    comboBox_CompanyName.BackColor = Color.Red;
            }
            catch { }

        }
        private void button_DataBasenameConfirm_Click(object sender, EventArgs e)
        {
            RE1.DataBaseFinder(comboBox_DataBaseName.Text, DatabaseList);
            comboBox_DataBaseName.DataSource = RE1.FilenamesForSearch;
            comboBox_DataBaseName.BackColor = Color.Red;
        }
      
        private void button_Refresh_Click(object sender, EventArgs e)
        {
           if(DC.Flag_searchDirectory)
            {
                myBackgroundWorker.RunWorkerAsync(2);
                button_Refresh.BackColor = Color.LightGreen;
                DC.Flag_searchDirectory = false; DC.Flag_forDisplayOfDatabase = true;
            }
        }

        private void monthCalendarEnd_Leave(object sender, EventArgs e)
        {
            monthCalendarEnd.Visible = false;
        }

        private void radioButton_LoraVision_CheckedChanged(object sender, EventArgs e)
        {
            
            if(radioButton_LoraVision.Checked)
            {
                comboBox_DataBaseName.Text = "AlsoEnergy2021Vision";
                if (!string.IsNullOrEmpty(comboBox_DataBaseName.Text))
                    comboBox_DataBaseName.BackColor = Color.LightGreen;
                else
                    comboBox_DataBaseName.BackColor = Color.Red;
            }
        }

        private void radioButton_Austin2020Vision_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton_Austin2020Vision.Checked)
            {
                comboBox_DataBaseName.Text = "Austin2021Vision";
                if (!string.IsNullOrEmpty(comboBox_DataBaseName.Text))
                    comboBox_DataBaseName.BackColor = Color.LightGreen;
                else
                    comboBox_DataBaseName.BackColor = Color.Red;
            }
            
        }

        private void comboBox_DataBaseName_DropDownClosed(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(comboBox_DataBaseName.Text))
            {
                comboBox_DataBaseName.BackColor = Color.LightGreen;
                comboBox_tab5_DBName.BackColor = Color.LightGreen;
            }
            else
            {
                comboBox_DataBaseName.BackColor = Color.Red;
                comboBox_tab5_DBName.BackColor = Color.Red;
            }
                
        }

        private void Button_main_MouseEnter(object sender, EventArgs e)
        {
            if(DC.Flag_forDisplayOfDatabase)
            {
                this.FileNameFromRootDir = string.Empty; this.DC.FilePathOfXML = string.Empty;
                this.DC.ExportXlSXPath = string.Empty; RE1.CompanyName = string.Empty; //IamPopUP = false;
                if (!string.IsNullOrEmpty(comboBox_CompanyName.Text))
                    this.FileNameFromRootDir = RE1.XMLFilePicker(comboBox_CompanyName.Text);//XMLFilePicker
                if (!string.Equals(FileNameFromRootDir, "ERROR SELECTION"))
                {
                    this.DC.FilePathOfXML = RE1.FilePathOfXMLtemp;
                    this.DC.ExportXlSXPath = RE1.ExportXlSXfilePath;
                    //comboBox_CompanyName.BackColor = Color.LightGreen;
                    RE1.CompanyName = comboBox_CompanyName.Text;
                    label_Database.Text = comboBox_CompanyName.Text; label_Database.Visible = true;
                    RTB1.AppendText("The Company you have selected: " + comboBox_CompanyName.Text+", DB: "+ comboBox_DataBaseName.Text);

                    if (DC.IamPopUP)
                    {
                        if (!string.IsNullOrEmpty(comboBox_CompanyName.Text) && !checkBox_SupressWarnings.Checked)
                            MessageBox.Show(comboBox_CompanyName.Text + " - is the Company name currently selected!\r\n" +
                                comboBox_DataBaseName.Text + " - is the DataBase.");
                        DC.IamPopUP = false;
                    }
                }
                DC.Flag_forDisplayOfDatabase = false;
            }
        }

        private void textBox_CustomerPO_MouseClick(object sender, MouseEventArgs e)
        {
            if(DC.flag_POnumber && !checkBox_SupressWarnings.Checked)
            {
                //MessageBox.Show("Important: The xls file naming does not support \"/,<space>\".");
                DC.flag_POnumber = false;
            }
        }

        private void comboBox_DataBaseName_DropDown(object sender, EventArgs e)
        {
            RE1.DataBaseFinder(comboBox_DataBaseName.Text, DatabaseList);
            comboBox_DataBaseName.DataSource = RE1.FilenamesForSearch; comboBox_tab5_DBName.DataSource = RE1.FilenamesForSearch;
            comboBox_DataBaseName.BackColor = Color.Red; comboBox_tab5_DBName.BackColor = Color.Red; labeltab5_1.Visible = false;
        }


        private void comboBox_CompanyName_DropDown(object sender, EventArgs e)
        {
            RE1.CompanyFinder(comboBox_CompanyName.Text);
            comboBox_CompanyName.DataSource = RE1.FilenamesForSearch;
            comboBox_CompanyName.BackColor = Color.Red;
            DC.Flag_forDisplayOfDatabase = true;
        }

        private void button_ForDebug_Click(object sender, EventArgs e)
        {
            textBox_PickTicketNumber.Text = "194371";
            comboBox_CompanyName.Text = "Fletcher-Reinhardt";
            comboBox_DataBaseName.Text = "Austin2020Vision";
            textBox_CustomerPO.Text = "D";
        }

        private void MachineLearning(string company, string PO, string DB, string Salesperson)
        {
            //\\netserver3\data\Log_Tickets_all\MachineLearning
        }

        #endregion Tab 1

        #region Tab 2

        private void button_CompanyCreation_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox_CreationCompanyname.Text) && !string.IsNullOrEmpty(textBox_FolderName.Text))
            {
                DC.Flag_searchDirectory = true;
                FormatModifier FP = new FormatModifier(richTextBox_CreationFormatForXML.Text);
                FP.FormatParser();
                FP.XMLCreator(DC.ROOTDIRFORXMLFILES_, textBox_CreationCompanyname.Text, textBox_FolderName.Text);
                Directory.CreateDirectory(DeclarationClass.SHIPMENTPATH_ + textBox_FolderName.Text);
                Directory.CreateDirectory(DeclarationClass.VISHALSHIPMENTPATH_ + textBox_FolderName.Text);

                if (checkBox_SupressWarnings.Checked)
                {
                    richTextBox_Tab2.Text = DeclarationClass.NotificationString3 +
                        "\r\n\r\nCompany Name: " + textBox_CreationCompanyname.Text +
                        "\r\nColumns It has created,";
                }
                else
                {
                    richTextBox_Tab2.Text = "Company Name: " + textBox_CreationCompanyname.Text +
                  "\r\nColumns It has created,";
                    MessageBox.Show(DeclarationClass.NotificationString3);//can be supressed.
                }

                //refresh the directories here.
                myBackgroundWorker.RunWorkerAsync(2);

                button_Refresh.BackColor = Color.LightGreen;
                DC.Flag_searchDirectory = false; DC.Flag_forDisplayOfDatabase = true;
            }
            else
                richTextBox_Tab2.Text = "Type all the necessary data required for Creation of Company and Company XML.\r\nTry again";

        }

        public void FormatSorter()
        {
            switch(comboBox_tab2_Typer.Text)
            {
                /*  LORA
                    ERT
                    CATM1
                    General
                */
                case "LORA":
                    richTextBox_CreationFormatForXML.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage,DevEUI,CommID,CommID1,CommID2,CommID3,CommID4," +
                         "ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
                    break;
                case "CATM1":
                    richTextBox_CreationFormatForXML.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage,AlternateID,IMEI," +
                        "SimCardID,ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
                    break;
                case "ERT":
                    richTextBox_CreationFormatForXML.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage," +
                       "CommID,8digitCommID,CommID1,CommID2,CommID3,CommID4,ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
                    break;
                case "General":
                    richTextBox_CreationFormatForXML.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage," +
                       "AlternateID,IMEI,SimCardID,DevEUI,CommID,8digitCommID,CommID1,CommID2,CommID3,CommID4,ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
                    break;
                default:
                    richTextBox_CreationFormatForXML.Text = "Company,PO#,Batch,FirmwareRevision,StatusCode,MeterID,KwhUsage," +
                       "AlternateID,IMEI,SimCardID,DevEUI,CommID,8digitCommID,CommID1,CommID2,CommID3,CommID4,ManufacturerType,MeterTypeCode,ClassAmps,Form/Base,ALSF,ALSL,ALSP,ALWA,Box,Pallet,Comments";
                    break;

            }
        }
  
        private void button4_Click(object sender, EventArgs e)
        {
            radioButton_Austin2020Vision.Checked = false;
            radioButton_LoraVision.Checked = false;

            textBox_PickTicketNumber.Text = string.Empty;
            comboBox_DataBaseName.Text = string.Empty;
            textBox_CustomerPO.Text = string.Empty;
            comboBox_CompanyName.Text = string.Empty;
            label_TicketNumberDisplay.Text = string.Empty;
            label_Database.Text = string.Empty;

            richTextBox_FileFormat.Clear();
            RTB1.Clear();

            comboBox_CompanyName.BackColor = Color.Red;
            comboBox_DataBaseName.BackColor = Color.Red;

        }

        private void comboBox_Tab2_CompanyName_DropDown(object sender, EventArgs e)
        {
            RE1.CompanyFinder(comboBox_Tab2_CompanyName.Text);
            comboBox_Tab2_CompanyName.DataSource = RE1.FilenamesForSearch;

        }

        private void richTextBox_FileFormat_TextChanged(object sender, EventArgs e)
        {
            FormatModifier FM = new FormatModifier(richTextBox_FileFormat.Text); XMLParser xm = new XMLParser();
            FM.FormatParser();
        }

        private void comboBox_tab2_Typer_DropDownClosed(object sender, EventArgs e)
        {
            richTextBox_CreationFormatForXML.Text = string.Empty;
            FormatSorter();
        }

        private void checkBox_tab1_Intellicode_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox_tab1_Intellicode.Checked)
            {
                checkBox_tab1_Intellicode.ForeColor = Color.Green;
            }
            if (!checkBox_tab1_Intellicode.Checked)
            {
                checkBox_tab1_Intellicode.ForeColor = Color.Red;
            }
        }

        private void checkBox_tab1_deleteEmptyCinXLS_CheckStateChanged(object sender, EventArgs e)
        {
            checkBox_tab1_deleteEmptyCinXLS.ForeColor = checkBox_tab1_deleteEmptyCinXLS.Checked ? Color.Red : Color.Green;
        }

        private void textBox_PickTicketNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            System.Windows.Forms.TextBox textbox = sender as System.Windows.Forms.TextBox;

            if (textbox == null)
                return;

            if (!char.IsControl(ch) &&  (!char.IsNumber(ch)) &&
                (ch != ',') && (ch != ',') && (ch != '.'))
                e.Handled = true;
        }

        private void button_Tab5_Click(object sender, EventArgs e)
        {
            try
            {
                PBar1.Maximum = 100; PBar1.Value = 0;
                labeltab5_1.Visible = true;
                labeltab5_1.Text = comboBox_tab5_DBName.Text;
                QueryTest QT = new QueryTest();
                QT.USER_init(comboBox_tab5_DBName.Text);

                PBar1.Value += 10;
                DC.databaseType = comboBox_tab5_DBName.Text.ToUpper().EndsWith("VISION") ? "dbo" : "power";
                dataGridView1.DataSource = QT.Tab5_AllDataQuery(textBox_tab5_PickTicket.Text, comboBox_tab5_DBName.Text, DC.databaseType);
                int ColumnCount = dataGridView1.ColumnCount;
            }
            catch
            {
                MessageBox.Show("Error Recalling the Data, Someting is missing!"); PBar1.Value = PBar1.Maximum;
            }
            PBar1.Value = PBar1.Maximum;
        }

        private void button_t6_Start_Click(object sender, EventArgs e)
        {
            #region ProgressBar
            PBar1.Minimum = 0; PBar1.Maximum = 200; PBar1.Value = 0;
            PBar1.Value += 50; label32.Visible = true;
            #endregion ProgressBar

            string XMLFormat = ".xml", XLSXFormat = ".xlsx", CSVFormat = ".csv", TEXTformat = ".txt";

            QueryTest QT = new QueryTest();//database query init

            DC.databaseType = comboBox_t6_DBName.Text.ToUpper().EndsWith("VISION") ? "dbo" : "power";

            string FileNameExtension = comboBox_t6_CompanyName.Text + "_PT" + textbox_t6_ticket.Text + "_PO" +
                           textbox_t6_PO.Text + "_" + GetCurrentDateAndTime(true);

            string CompletePathForExport = DC.ORIGINALSHIPMENTPATH_ + comboBox_t6_CompanyName.Text+ @"\" + FileNameExtension;

            richTextBox_T6.Text = "File path created.\r\n";

            PBar1.Value += 20;

            QT.USER_init(comboBox_t6_DBName.Text); //dataquery user init

            QT.TestQuerySpcl(Spcl_DatDBColumnNames, Spcl_FileColumnNames, Spcl_ValueForColumnStatics, Spcl_MergeEvents, textbox_t6_ticket.Text, DC.databaseType , WhatToFind.Text, textbox_t6_PO.Text);
            this.Spcl_ArrayMessageFromDatabase = QT.ArrayMessageFromDatabase;

            PBar1.Value += 20;
            richTextBox_T6.AppendText("Datasbase access complete.\r\n");
            ExcelProcessor EXS = new ExcelProcessor();
            /*
             *This function actually check weather we opt for csv or xls.
             */
            if(radioButton1.Checked)   //csv file creation  && !radioButton2.Checked
            {
                CompletePathForExport += CSVFormat;
                richTextBox_T6.AppendText("\r\nFile is being written.");
                EXS.WriteCSVSpecial(Spcl_FileColumnNames, QT.RowCounter, Spcl_ArrayMessageFromDatabase, CompletePathForExport);
            }

            else if (radioButton_T6_TXT.Checked)   //csv file creation  && !radioButton2.Checked
            {
                CompletePathForExport += TEXTformat;
                richTextBox_T6.AppendText("\r\nFile is being written.");
                EXS.WriteCSVSpecial(Spcl_FileColumnNames, QT.RowCounter, Spcl_ArrayMessageFromDatabase, CompletePathForExport);
            }


            else if(radioButton2.Checked)  //excel File creation !radioButton1.Checked && 
            {
                CompletePathForExport += XLSXFormat;
                richTextBox_T6.AppendText("\r\nFile is being written.");
                EXS.WriteExcelSpecial(Spcl_FileColumnNames, QT.RowCounter, Spcl_ArrayMessageFromDatabase, CompletePathForExport);
            }

            else if(radioButton3_XML.Checked) //XML format
            {
                CompletePathForExport += XMLFormat;
                richTextBox_T6.AppendText("\r\nFile is being written.");
                EXS.WriteXMLSpecial(Spcl_FileColumnNames, QT.RowCounter, Spcl_ArrayMessageFromDatabase,Place1,Place2,Place3,Place4, textBox_DeviceType.Text, CompletePathForExport);
            }
            richTextBox_T6.AppendText("\r\nWriting is Done");
            PBar1.Value += 20;

            MBU.MB_TextDisplay("The File is created.\r\n"+ CompletePathForExport); button_t6_browse.BackColor = Color.Transparent;
            PBar1.Value = PBar1.Maximum;

            DialogResult dialogR = MBU.ShowDialog();

            DataLogging DLT = new DataLogging();
            try 
            {
                DC.Log_DataCollectionString = RTB1.Text;

                DC.Log_TicketToLog += "<" + textbox_t6_ticket.Text + "> ";

                try { DLT.FileOpener(DC.Log_TicketToLog, "S", "-------------------------------------------------"); }
                catch { richTextBox_T6.AppendText("\r\nError in the data Logging"); if (!checkBox_SupressWarnings.Checked) { MessageBox.Show("Important-- Log is not recorded due to some Error."); } }

            }
            catch
            {
                richTextBox_T6.AppendText("\r\nError in the data Logging");
            }
        }

        private void comboBox_t6_CompanyName_DropDown(object sender, EventArgs e)
        {
            RE1.CompanyFinder(comboBox_t6_CompanyName.Text);
            comboBox_t6_CompanyName.DataSource = RE1.FilenamesForSearch;
            comboBox_t6_CompanyName.BackColor = Color.Red;
            DC.Flag_forDisplayOfDatabase = true;
        }

        private void comboBox_t6_CompanyName_DropDownClosed(object sender, EventArgs e)
        {

            try
            {
                if (!string.IsNullOrEmpty(comboBox_t6_CompanyName.Text))
                {
                    comboBox_t6_CompanyName.BackColor = Color.LightGreen;
                    DC.Flag_forDisplayOfDatabase = true;
                }
                else
                    comboBox_t6_CompanyName.BackColor = Color.Red;
            }
            catch { }
        }

        private void comboBox_t6_DBName_DropDown(object sender, EventArgs e)
        {
            RE1.DataBaseFinder(comboBox_t6_DBName.Text, DatabaseList);
            comboBox_t6_DBName.DataSource = RE1.FilenamesForSearch; comboBox_t6_DBName.DataSource = RE1.FilenamesForSearch;
            comboBox_t6_DBName.BackColor = Color.Red; comboBox_t6_DBName.BackColor = Color.Red;
        }

        private void comboBox_t6_DBName_DropDownClosed(object sender, EventArgs e)
        {
            comboBox_t6_DBName.BackColor = string.IsNullOrEmpty(comboBox_t6_DBName.Text) ? Color.Red : Color.LightGreen;
        }

        private void checkBox_tab4_Everything_CheckStateChanged(object sender, EventArgs e)
        {
            if(checkBox_tab4_Everything.Checked)
            {
                label18.Visible = false;
                textBox_tab4_month.Visible = false;
                monthCalendar_tab4.Visible = false;
            }
            if (!checkBox_tab4_Everything.Checked)
            {
                label18.Visible = true;
                textBox_tab4_month.Visible = true;
                monthCalendar_tab4.Visible = true;
            }

        }

        private void comboBox_tab5_DBName_DropDown(object sender, EventArgs e)
        {
            RE1.DataBaseFinder(comboBox_tab5_DBName.Text, DatabaseList); comboBox_tab5_DBName.DataSource = RE1.FilenamesForSearch;
           comboBox_tab5_DBName.BackColor = Color.Red; labeltab5_1.Visible = false;
        }

        private void comboBox_tab5_DBName_DropDownClosed(object sender, EventArgs e)
        {
            comboBox_tab5_DBName.BackColor = string.IsNullOrEmpty(comboBox_DataBaseName.Text) ? Color.LightGreen : Color.Red; comboBox_DataBaseName.Text = comboBox_tab5_DBName.Text;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                radioButton2.Checked = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                radioButton1.Checked = false;
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            richTextBox_5.Visible = checkBox1.Checked ? true : false;
        }

        private void button5_Click(object sender, EventArgs e)//get data 2
        {
            try
            {
                PBar1.Maximum = 100; PBar1.Value = 0;
                labeltab5_1.Visible = true;
                labeltab5_1.Text = comboBox_tab5_DBName.Text;
                QueryTest QT = new QueryTest();
                QT.USER_init(comboBox_tab5_DBName.Text);

                PBar1.Value += 10;

                DC.databaseType = comboBox_tab5_DBName.Text.ToUpper().EndsWith("VISION") ? "dbo" : "power";

                AryOfColumns = QT.Tab5_ColumnNameQuery(textBox_tab5_PickTicket.Text, comboBox_tab5_DBName.Text, "street" , DC.databaseType);
                richTextBox_5.Clear();
                foreach(string ColumnHead in AryOfColumns)
                {
                    richTextBox_5.AppendText(ColumnHead+"\r\n");
                }
            }
            catch
            {
                MessageBox.Show("Error Recalling the Data, Someting is missing!"); PBar1.Value = PBar1.Maximum;
            }
            PBar1.Value = PBar1.Maximum;
            if(AryOfColumns.Count()>1)
            {
                textBoxT5_SearchTB.Visible = true;
                buttonT5_Search.Visible = true;
            }
        }

        private void buttonT5_Search_Click(object sender, EventArgs e)
        {
            try
            {
                if (AryOfColumns.Count() > 1)
                {
                    int demo = AryOfColumns.Length;int count = 0;
                    richTextBox_5.Clear();
                    richTextBox_5.Text = "Results for " + textBoxT5_SearchTB.Text + " are being shown:\r\n\r\n";
                    foreach (string ColumnHead in AryOfColumns)
                    {
                        if(ColumnHead!=null)
                        {
                            if (ColumnHead.ToUpper().Contains(textBoxT5_SearchTB.Text.ToUpper()))
                            {
                                richTextBox_5.AppendText(ColumnHead + "        ......Position: "+count+"\r\n");
                            }
                        }
                        else
                        {
                            richTextBox_5.AppendText("Total number of columns in the Table are Approx. :: "+count+"\r\n");
                            break;
                        }
                            
                        count++;
                    }
                }
            }
            catch{}
        }

        private void button8_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                try
                {
                    openFileDialog1.InitialDirectory = DC.folderNameForOutputFile;//@"\\netserver3\DATA";
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls";//|All files (*.*)|*.*
                    openFileDialog1.FilterIndex = 2;
                    openFileDialog1.RestoreDirectory = true;
                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    { }
                }
                catch { }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.Clear();
                Clipboard.SetText(DC.FileNameExtension_Global);
            }
            catch { }
        }

        private void button_tab5_GetDB_Click(object sender, EventArgs e)
        {
            richTextBox_5.Clear(); PBar1.Value = 0; PBar1.Maximum = DatabaseList.Count;
            QueryTest SQL = new QueryTest();
            foreach(string DBelement in DatabaseList)
            {
                if(DBelement.ToUpper().EndsWith("VISION"))
                {
                    PBar1.Value += 1;
                    string tempBatchID = SQL.FindTheDBwithMeterID(textBox_tab5_TicketToSearch.Text, DBelement, comboBox_DBOtype.Text);
                    if (!string.Equals(tempBatchID, "NoData"))
                    {
                        //textBox_tab5_DisplayDB.Text = DBelement;
                        richTextBox_5.Text = "The meterID " + textBox_tab5_TicketToSearch.Text + " is located in " + DBelement + " with BatchID: " + tempBatchID + "\r\n";
                        break;
                    }
                }
            }
            PBar1.Value = PBar1.Maximum;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string commandSQL = richTextBox_5.Text;
            QueryTest SQL = new QueryTest();
            dataGridView1.DataSource = SQL.SendSQLRaw(commandSQL, comboBox_dB_SQL_RAW.Text);
        }

        private void comboBox_dB_SQL_RAW_DropDown(object sender, EventArgs e)
        {
            RE1.DataBaseFinder(comboBox_dB_SQL_RAW.Text, DatabaseList); comboBox_dB_SQL_RAW.DataSource = RE1.FilenamesForSearch;
            comboBox_dB_SQL_RAW.BackColor = Color.Red; labeltab5_1.Visible = false;
        }

        private void comboBox_dB_SQL_RAW_DropDownClosed(object sender, EventArgs e)
        {
            comboBox_dB_SQL_RAW.BackColor = string.IsNullOrEmpty(comboBox_dB_SQL_RAW.Text) ? Color.LightGreen : Color.Red;
        }

        private void button_t6_browse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
            {
                openFileDialog1.InitialDirectory = @"\\netserver3\data\ShipmentsXMLfiles\ExcelFormatForSpclFiles";//@"\\netserver3\DATA";
                openFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    textBox_t6_excellFilePath.Text = openFileDialog1.FileName;
                    button_t6_Start.Visible = false;
                    myBackgroundWorkertab6.RunWorkerAsync(2);
                }
                else if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                {
                    textBox_t6_excellFilePath.Text = "Nothing Selected";
                }
                if(textBox_t6_excellFilePath.Text.ToLower().Contains("entergy"))
                {
                    richTextBox_T6.AppendText("Entergy has a important need of PO number. Be specific!\r\nCSV file as output.");
                }
                else if(textBox_t6_excellFilePath.Text.ToLower().Contains("dominion"))
                {
                    richTextBox_T6.AppendText("Dominion needs xml format and Device type, Please mention before generating output!\r\nXML as output");
                }
            }
        }

        private void textBox_CustomerPO_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            System.Windows.Forms.TextBox textbox = sender as System.Windows.Forms.TextBox;

            if (textbox == null)
                return;

            if (!char.IsControl(ch) && (!char.IsNumber(ch)) && !char.IsLetterOrDigit(ch) &&
                (ch != '_'))
                e.Handled = true;
        }

        private void comboBox_Tab2_CompanyName_DropDownClosed(object sender, EventArgs e)
        {
            textBox_CreationCompanyname.Text = comboBox_Tab2_CompanyName.Text;
            DC.Flag_InsertNFHere = false;
        }

        private void button_PasteCompanyName_Click(object sender, EventArgs e)
        {
            if(!DC.Flag_InsertNFHere)
            {
                textBox_FolderName.Text = textBox_CreationCompanyname.Text; DC.Flag_InsertNFHere = true;
            }
            else
            {
                textBox_FolderName.Text = textBox_CreationCompanyname.Text + "_NF"; DC.Flag_InsertNFHere = true;
            }
        }

        private void textBox_PickTicketNumber_Click(object sender, EventArgs e)
        {
            if(DC.Flag_searchDirectory)
            {
                myBackgroundWorker.RunWorkerAsync(2);
                DC.Flag_searchDirectory = false;
            }
        }
        
       

        #endregion Tab 2

        #region tab 3

        private void buttonTab3_Start(object sender, EventArgs e)
        {
            Authentication AU = new Authentication();
            DialogResult dialogR = AU.ShowDialog();
            if(dialogR == DialogResult.OK)
            {
                if (AU.ChecktheUserPass())
                {
                    richTextBox_TAB3.Text = "The Credentials are Correct.";
                    AU.Dispose();
                    if (!string.IsNullOrEmpty(textBox_FolderPath.Text) && !string.IsNullOrEmpty(textBox_StartDate.Text))
                    {
                        richTextBox_TAB3.Text = "The Credentials are Correct." +
                        "\r\nThe Process is started!";
                        Button_main.Visible = false;
                        myBackgroundWorkerTab3.RunWorkerAsync(2);
                        DC.Flag_searchDirectory = true;
                    }
                    else
                    {
                        richTextBox_TAB3.Text = "The Credentials are Correct." +
                      "\r\nThe Process is halted as The entries are not sufficient!";
                    }
                }
                else
                {
                    richTextBox_TAB3.Text = "You have not entered the Credentials correctly." +
                          "\r\nThe Process is halted." +
                          "\r\nHint: its PinCode";
                }
            }
            
            if(dialogR == DialogResult.Cancel)
            {
                richTextBox_TAB3.Text = "You have not entered the Credentials." +
                        "\r\nThe Process is halted.";
            }
        }
        private void textBox3_StartDate_MouseClick(object sender, MouseEventArgs e)
        {
            monthCalendarStart.Visible = true; monthCalendarEnd.Visible = false;
        }

        private void monthCalendarStart_DateSelected(object sender, DateRangeEventArgs e)
        {
            textBox_StartDate.Text = monthCalendarStart.SelectionRange.Start.ToString("MM/dd/yyyy");
            this.StartDate = monthCalendarStart.SelectionRange.Start;
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void textBox3_EndDate_MouseClick(object sender, MouseEventArgs e)
        {
            monthCalendarEnd.Visible = true; monthCalendarStart.Visible = false;
        }

        private void monthCalendarEnd_DateSelected(object sender, DateRangeEventArgs e)
        {
            textBox_EndDate.Text = monthCalendarEnd.SelectionRange.Start.ToString("MM/dd/yyyy");
            this.EndDate = monthCalendarEnd.SelectionRange.Start;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var openFileDialog1 = new FolderBrowserDialog())
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    DC.XMLMakerPath = openFileDialog1.SelectedPath;
                    textBox_FolderPath.Text = DC.XMLMakerPath;
                }
            }
        }
        #endregion tab 3

        #region tab 4

        private void monthCalendar_tab4_DateSelected(object sender, DateRangeEventArgs e)
        {
            richTextBox_tab4.Clear();
            textBox_tab4_month.Text = monthCalendar_tab4.SelectionStart.ToString("MMMM dd yyyy");
            DC.YearForSearch = monthCalendar_tab4.SelectionStart.ToString("yyyy");
            DC.Search_TicketNumber = textBox_tab4_SearchTicket.Text;
            switch (int.Parse(monthCalendar_tab4.SelectionStart.ToString("MM")))
            {
                case 01:
                    DC.month_T = "January";
                    DC.month_Tminus1 = "December";
                    DC.month_Tplus1 = "February";
                    break;

                case 02:
                    DC.month_T = "February";
                    DC.month_Tminus1 = "January";
                    DC.month_Tplus1 = "March";
                    break;

                case 03:
                    DC.month_T = "March";
                    DC.month_Tminus1 = "February";
                    DC.month_Tplus1 = "April";
                    break;

                case 04:
                    DC.month_T = "April";
                    DC.month_Tminus1 = "March";
                    DC.month_Tplus1 = "May";
                    break;

                case 05:
                    DC.month_T = "May";
                    DC.month_Tminus1 = "April";
                    DC.month_Tplus1 = "June";
                    break;

                case 06:
                    DC.month_T = "June";
                    DC.month_Tminus1 = "May";
                    DC.month_Tplus1 = "July";
                    break;

                case 07:
                    DC.month_T = "July";
                    DC.month_Tminus1 = "June";
                    DC.month_Tplus1 = "August";
                    break;

                case 08:
                    DC.month_T = "August";
                    DC.month_Tminus1 = "July";
                    DC.month_Tplus1 = "September";
                    break;

                case 09:
                    DC.month_T = "September";
                    DC.month_Tminus1 = "August";
                    DC.month_Tplus1 = "October";
                    break;

                case 10:
                    DC.month_T = "October";
                    DC.month_Tminus1 = "September";
                    DC.month_Tplus1 = "November";
                    break;

                case 11:
                    DC.month_T = "November";
                    DC.month_Tminus1 = "October";
                    DC.month_Tplus1 = "December";
                    break;

                case 12:
                    DC.month_T = "December";
                    DC.month_Tminus1 = "November";
                    DC.month_Tplus1 = "January";
                    break;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            richTextBox_tab4.Clear();
            richTextBox_tab4.AppendText("\r\nThe Ticket is being searched!\r\n");
            myBackgroundWorkerTab4.RunWorkerAsync(2);
        }

        #endregion tab 4

        #region Common functions
        public void Ticket_Formater(string FormatString, string dbo_type)
        {
            if (FormatString.Contains(",") || FormatString.Contains(", ") || FormatString.Contains(" ,") || FormatString.Contains(" , ") ||
                FormatString.Contains(".") || FormatString.Contains(". ") || FormatString.Contains(" .") || FormatString.Contains(" . "))
            {
                int StartIndex = 0, StopIndex = 0; bool LocalFlag_ForLoop = true;

                string TempString = FormatString.Replace(", ", ",");
                TempString = TempString.Replace("\n", string.Empty); TempString = FormatString.Replace(" , ", ","); TempString = FormatString.Replace(" ,", ",");
                TempString = TempString.Trim(',', '.');
                int LenghtOfTempString = TempString.Length;

                while (LocalFlag_ForLoop)
                {
                    try
                    {
                        try
                        {
                            StopIndex = TempString.IndexOf(',');
                            if (StopIndex == -1 || StopIndex < 0)
                            {
                                StopIndex = TempString.Length; LocalFlag_ForLoop = false;
                                TicketsListForDataQuerySQL.Add("(((" + dbo_type + ".Meter.Batch)='" + TempString + "'))");
                                TicketNumberIndividual.Add(TempString);
                                break;
                            }
                        }
                        catch { StopIndex = TempString.Length; LocalFlag_ForLoop = false; }

                        TicketsListForDataQuerySQL.Add("(((" + dbo_type + ".Meter.Batch)='" + TempString.Substring(StartIndex, StopIndex) + "'))");
                        TicketNumberIndividual.Add(TempString.Substring(StartIndex, StopIndex));

                        TempString = TempString.Substring(StopIndex + 1, TempString.Length - (StopIndex + 1));
                        StartIndex = 0;
                    }
                    catch (Exception ex)
                    {
                        if(!checkBox_SupressWarnings.Checked)
                            MessageBox.Show("Error in the formatParser ticket function\r\nStopIndex < 0 can be a error\r\n" + ex);
                    }
                }
            }
            else
            {
                TicketsListForDataQuerySQL.Add("((("+ dbo_type + ".Meter.Batch)='" + FormatString + "'))");
                TicketNumberIndividual.Add(FormatString);
            }
        }
        public string GetCurrentDateAndTime(bool forFile)
        {
            string dateFormatYY_MM_dd;
            DateTime lastupdated = DateTime.Today;
            if (forFile)
            {
                dateFormatYY_MM_dd = lastupdated.ToString("yyyyMMdd");
                return dateFormatYY_MM_dd;
            }
            else
            {
                dateFormatYY_MM_dd = lastupdated.ToString("yyyy/MM/dd");
                return dateFormatYY_MM_dd;
            }
        }

        public string SalespersonQuiz(string comboBox_SalesPerson)
        {
            if (string.Equals(comboBox_SalesPerson, "JLH"))
                 return "Jessica@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "MJM"))
                 return "maria@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "JDD"))
                 return "jesse@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "DDR"))
                 return "debbie@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "RHA"))
                 return "randy@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "SEI"))
                 return "samantha@visionmetering.com";
            if (string.Equals(comboBox_SalesPerson, "TRN"))
                 return "tom@visionmetering.com";
            else
                return "No Correct Salesperson Selected.";

        }
        /*the function  helps to verify the username and password to make the changes to the application
         */
        #region Authentication
        public string CheckTheUsername(string INPUT)
        {
            foreach(string user in users)
            {
                if(string.Equals(INPUT.ToUpper(),user.ToUpper()))
                    return "OK";
            }
            return "ERROR";
        }
        #endregion Authentication

        #endregion Common functions

        #endregion Adjoining Functions

        #region Threading


        private BackgroundWorker myBackgroundWorker;//myBackgroundWorker.RunWorkerAsync(2)
        #region myBackgroundWorker
        private void myBackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            worker.ReportProgress(30);
            QueryTest DBq = new QueryTest();
            System.Data.DataTable dt = DBq.GetDataTables();
            worker.ReportProgress(20);
            for (int counter = 0; counter < dt.Rows.Count; counter++)
            {
                DatabaseList.Add(dt.Rows[counter][0].ToString());
                worker.ReportProgress(counter);
            }
            worker.ReportProgress(20);
            button_Refresh.BackColor = Color.LightGreen;
            RE1.DirectoriesExplorer();//form loads slowly
            worker.ReportProgress(20);
            DC.Flag_searchDirectory = false;
           
        }

        private void myBackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.PBar1.Value = this.PBar1.Maximum; button_Refresh.Visible = true; //button invisible for a while
        }

        private void myBackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try { this.PBar1.Value = e.ProgressPercentage; }
            catch { this.PBar1.Refresh(); }
            this.PBar1.Refresh();
        }
        #endregion #region myBackgroundWorker

        private BackgroundWorker myBackgroundWorkerTab3;//myBackgroundWorkerTab3.RunWorkerAsync(2)
        #region myBackgroundWorkerTab3
        private void myBackgroundWorkerTab3_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker1 = sender as BackgroundWorker;
            if (!string.IsNullOrEmpty(textBox_FolderPath.Text) && !string.IsNullOrEmpty(textBox_StartDate.Text))
            {
                worker1.ReportProgress(20);
                DC.Flag_searchDirectory = true;//flag to update the Directory as you type the ticket number
                RootDirectoriesExplorer RE0 = new RootDirectoriesExplorer(); ExcelProcessor EX0 = new ExcelProcessor();
                RE0.FileExplorerForXML(textBox_FolderPath.Text, StartDate, EndDate);
                counterForFileGeneratedInXml = EX0.ExcelExtraction(RE0.FileNames, RE0.FileDirecrtory, RE0.DirNames, textBox_FolderPath.Text, DC.PARENTFOLDERTOSTICKTO_);
                worker1.ReportProgress(20);
            }
        }

        private void myBackgroundWorkerTab3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.PBar1.Value = this.PBar1.Maximum;
            richTextBox_TAB3.AppendText("\r\n"+ counterForFileGeneratedInXml +"-- Files are exported."); Button_main.Visible = true;
        }

        private void myBackgroundWorkerTab3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try { this.PBar1.Value = e.ProgressPercentage; }
            catch { this.PBar1.Refresh(); }
            richTextBox_TAB3.AppendText("\r\nprogress is going on!");
            this.PBar1.Refresh();
        }
        #endregion myBackgroundWorkerTab3

        private BackgroundWorker myBackgroundWorkerTab4;//myBackgroundWorker.RunWorkerAsync(2)
        #region myBackgroundWorker
        private void myBackgroundWorkerTab4_DoWork(object sender, DoWorkEventArgs e)
        {
            string rootDLocal = @"\\netserver3\data\Log_Tickets_all";
            int ArrayCount = 0; Array.Clear(Dte, 0, Dte.Length);
            if (!checkBox_tab4_Everything.Checked)
            {
                string month = string.Empty;
                SearchDataTab4.Clear();
                BackgroundWorker worker4 = sender as BackgroundWorker;
                //RootDirectoriesExplorer RDE = new RootDirectoriesExplorer();
                //RDE.DirectoriesExplorer("\\netserver3\\data\\Log_Tickets_all", "*.txt");
                int counter = 1;
                while (SearchDataTab4.Count == 0 && counter <= 3)
                {
                    if (counter == 1)
                    {
                        month = DC.month_T;
                    }

                    if (counter == 2)
                    {
                        month = DC.month_Tminus1;
                        if (month.ToUpper().Contains("DEC"))
                        {
                            DC.YearForSearch = (int.Parse(DC.YearForSearch) - 1).ToString();
                        }
                    }
                    if (counter == 3)
                    {
                        month = DC.month_Tplus1;
                        if (month.ToUpper().Contains("JAN"))
                        {
                            DC.YearForSearch = (int.Parse(DC.YearForSearch) + 1).ToString();
                        }
                    }


                    string tempPATH = "\\\\netserver3\\data\\Log_Tickets_all\\TicketLog" + month + DC.YearForSearch + ".txt";
                    if (File.Exists(tempPATH))
                    {
                        DC.String_SearchDataTab4 = File.ReadAllText(tempPATH);
                        //string DemoTicketNumber = "<Ticket> <" + textBox_tab4_SearchTicket.Text + ">  </Ticket>";
                        string DemoTicketNumber = "<" + textBox_tab4_SearchTicket.Text + ">";
                        bool demo = DC.String_SearchDataTab4.Contains(DemoTicketNumber);//<Ticket> 193416 </Ticket>
                        string RemainingString;
                        if (demo)
                        {
                            do
                            {

                                DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf(DemoTicketNumber));
                                Dte[ArrayCount] = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf("<Date>") + 6, 12);//11/03/2020
                                RemainingString = DC.String_SearchDataTab4.Substring(DemoTicketNumber.Length);
                                int demoint = DC.String_SearchDataTab4.IndexOf("</Log>");
                                DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf("<Log>"));//, String_SearchDataTab4.IndexOf("</Log>"));
                                demoint = DC.String_SearchDataTab4.IndexOf("</Log>");
                                DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(5, DC.String_SearchDataTab4.IndexOf("</Log>") - 5);
                                SearchDataTab4.Add(DC.String_SearchDataTab4);
                                worker4.ReportProgress(10);
                                DC.String_SearchDataTab4 = string.Empty;
                                DC.String_SearchDataTab4 = RemainingString;
                                ArrayCount++;
                            }
                            while (RemainingString.Contains("<Ticket> <" + textBox_tab4_SearchTicket.Text + ">  </Ticket>"));
                        }

                        /*I am working here on the loop of do while to get as many results as mentioned in the log file.
                         * first result is ok but the other result is not getting as requsted.
                         */
                    }
                    counter++;
                }
            }
            else
            {
                try
                {
                   var dirs = from dir in
                   Directory.GetFiles(rootDLocal) //          EnumerateDirectories(rootDLocal)
                   select dir;

                    foreach (var dir in dirs)
                    {
                        if (File.Exists(dir))
                        {
                            DC.String_SearchDataTab4 = File.ReadAllText(dir);
                            //string DemoTicketNumber = "<Ticket> <" + textBox_tab4_SearchTicket.Text + ">  </Ticket>";
                            string DemoTicketNumber = "<" + textBox_tab4_SearchTicket.Text + ">";
                            bool demo = DC.String_SearchDataTab4.Contains(DemoTicketNumber);//<Ticket> 193416 </Ticket>
                            string RemainingString;
                            if (demo)
                            {
                                do
                                {
                                    DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf(DemoTicketNumber));
                                    Dte[ArrayCount] = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf("<Date>") + 6, 12);//11/03/2020
                                    RemainingString = DC.String_SearchDataTab4.Substring(DemoTicketNumber.Length);
                                    int demoint = DC.String_SearchDataTab4.IndexOf("</Log>");
                                    DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(DC.String_SearchDataTab4.IndexOf("<Log>"));//, String_SearchDataTab4.IndexOf("</Log>"));
                                    demoint = DC.String_SearchDataTab4.IndexOf("</Log>");
                                    DC.String_SearchDataTab4 = DC.String_SearchDataTab4.Substring(5, DC.String_SearchDataTab4.IndexOf("</Log>") - 5);
                                    SearchDataTab4.Add(DC.String_SearchDataTab4);
                                    DC.String_SearchDataTab4 = string.Empty;
                                    DC.String_SearchDataTab4 = RemainingString;
                                    ArrayCount++;
                                }
                                while (RemainingString.Contains("<Ticket> <" + textBox_tab4_SearchTicket.Text + ">  </Ticket>"));
                            }

                            /*I am working here on the loop of do while to get as many results as mentioned in the log file.
                             * first result is ok but the other result is not getting as requsted.
                             */
                        }
                    }
                }
                catch { }
            }


        }

        private void myBackgroundWorkerTab4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            int counter = 1;
            if (SearchDataTab4.Count != 0)
            {
                foreach (string n in SearchDataTab4)
                {
                    richTextBox_tab4.AppendText("\r\nResult no: " + counter); richTextBox_tab4.AppendText("\r\nDOL: " + Dte[counter-1] +"\r\n");
                    richTextBox_tab4.AppendText(n); richTextBox_tab4.AppendText("\r\n");
                    counter++;
                }
            }
            else
            {
                richTextBox_tab4.Text = "The search was unsuccessful! try Different ticket." +
                    "\r\nPossiblity of logging a wrong ticket is almost 0, Check for typo." +
                    "\r\nMonths searched for File - " + DC.month_Tminus1 + ", " + DC.month_T + ", " + DC.month_Tplus1;
            }
            SearchDataTab4.Clear();
        }

        private void myBackgroundWorkerTab4_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //richTextBox_tab4.Text = "\r\nThe search is a success:" + textBox_tab4_SearchTicket.Text + "\r\n" + String_SearchDataTab4;
        }
        #endregion #region myBackgroundWorker

        private BackgroundWorker myBackgroundWorkertab6;
        #region myBackgroundWorkerTab6
        private void myBackgroundWorkertab6_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker1 = sender as BackgroundWorker;

            ExcelProcessor EXLSPCL = new ExcelProcessor();
            EXLSPCL.PreprocessExcelSpcl(textBox_t6_excellFilePath.Text);
            this.Spcl_DatDBColumnNames = EXLSPCL.DatDBColumnNames;
            this.Spcl_FileColumnNames = EXLSPCL.FileColumnNames;
            this.Spcl_ValueForColumnStatics = EXLSPCL.ValueForColumnStatics;
            this.Spcl_MergeEvents = EXLSPCL.MergeEvents;

            this.Place1 = EXLSPCL.Place1;
            this.Place2 = EXLSPCL.Place2;
            this.Place3 = EXLSPCL.Place3;
            this.Place4 = EXLSPCL.Place4;

            this.Spcl_DatDBColumnNames.RemoveAt(0);
            this.Spcl_FileColumnNames.RemoveAt(0);
            this.Spcl_ValueForColumnStatics.RemoveAt(0);
            this.Spcl_MergeEvents.RemoveAt(0);

            try
            {
                this.Place1.RemoveAt(0);
                this.Place2.RemoveAt(0);
                this.Place3.RemoveAt(0);
                this.Place4.RemoveAt(0);
            }
            catch { }
        }

        private void myBackgroundWorkertab6_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button_t6_browse.BackColor = Color.Green; button_t6_Start.Visible = true;
        }

        private void myBackgroundWorkertab6_ProgressChanged(object sender, ProgressChangedEventArgs e){}
        #endregion myBackgroundWorkerTab6

        #endregion threading

        #region commented code
        //private void XlsCompareToAdd(List<string> TemoList1, List<string> TemoList2)
        //{
        //    for (int reference1 =0,reference2=0;reference1 < TemoList2.Count ||reference2 < TemoList1.Count; reference1++,reference2++)//int reference = TemoList1.Count - 1; reference >= 0; reference--
        //    {
        //        if(reference1< TemoList2.Count)
        //        {
        //            if (!TemoList1.Contains(TemoList2[reference1]))
        //            {
        //                ColumnNumberToAddFromFile2.Add(reference1 + 1);
        //                ColumnNameToAddFromFile2.Add(TemoList2[reference1]);
        //            }
        //        }

        //        if(reference2 < TemoList1.Count)//checks the condition
        //        {
        //            if (!TemoList2.Contains(TemoList1[reference2]))
        //            {
        //                ColumnNumberToDeleteFromFile1.Add(reference2+1);
        //                ColumnNameToDeleteFromFile1.Add(TemoList1[reference2]);
        //                //if (TemoList1[reference] == TemoList2[reference])
        //                //{

        //                //}
        //                //ColumnNumberToDeleteFromFile1.Add(reference);
        //                //ColumnNameToDeleteFromFile1.Add(TemoList2[reference])
        //            }
        //        }
        //        ////int numOfDuplicates = 1;
        //        //for (int comparingTo = TemoList2.Count - 2; comparingTo >= 0; comparingTo--)
        //        //{
        //        //    if (TemoList1[reference] == TemoList2[comparingTo])
        //        //        ColumnsToKeep++;

        //        //    else if(TemoList1[reference]!=TemoList2[comparingTo])
        //        //        ColumnNumberToDelete.Add(reference);
        //        //        ColumnNameToDelete.Add(TemoList2[reference]);
        //        //}
        //    }

        //}
        #endregion commented code
        #region commented code
        //public void DisplayText(string textContent,int Condition)
        //{
        //    switch(Condition)
        //    {
        //        case APPEND:
        //            richTextBox1.AppendText("\r\n" + textContent);
        //            break;
        //        case NEWLine:
        //            richTextBox1.Text = textContent;
        //            break;

        //        default:
        //            richTextBox1.Text = textContent;
        //            break;
        //    }

        //    //if(Condition.Contains("Append"))
        //    //{
        //    //    richTextBox1.AppendText("\r\n"+textContent);
        //    //}
        //    //else if(Condition.Contains("New"))
        //    //{
        //    //    richTextBox1.Text = textContent;
        //    //}
        //    //else
        //    //{
        //    //    richTextBox1.Text = textContent;
        //    //}
        //}

        //public static class Utilities
        //{
        //    #region Check For Null String

        //    public static string CheckForNullString(dynamic s)
        //    {
        //        string str = (s == null) ? string.Empty : s;
        //        return str;
        //    }

        //    #endregion Check For Null String
        //}

        //private void USER_init()  //commented fro testing only
        //{
        //    this.user = new User();
        //    user.Server = "Netserver3";
        //    user.Database = comboBox_DataBaseName.Text; //"Austin2020Vision";
        //    //user.Database = "LoraVision";
        //    user.DBOwner = "dbo";

        //    user.SQLCredentials = new Credentials();
        //    user.SQLCredentials.UserID = "power";
        //    user.SQLCredentials.Password = "power";

        //    user.SetConnectionString();  //connection string is set here
        //}

        //private void Tab1_TestQuery(List<string> Columnnames, string CompanyName)
        //{
        //    string batch = textBox_PickTicketNumber.Text;//"192735";
        //    string query =
        //        "SELECT * " +
        //        "FROM ((dbo.Meter INNER JOIN dbo.MeterTypeView ON dbo.Meter.MeterTypeCode = dbo.MeterTypeView.MeterTypeCode) " +
        //        "INNER JOIN dbo.MeterTest ON dbo.Meter.MeterID = dbo.MeterTest.MeterID) " +
        //        "INNER JOIN dbo.MeterReadings ON dbo.Meter.MeterID = dbo.MeterReadings.MeterID " +
        //        "WHERE (((dbo.Meter.Batch)='" + batch + "')) " +
        //        "ORDER BY dbo.Meter.MeterID, dbo.Meter.Box, dbo.Meter.Pallet,dbo.Meter.IMEI";

        //    try
        //    {
        //        this.dt = DatabaseQueries.ExecuteQuery(query, user.ConnectionString);
        //        //dynamic FormBAse;
        //        if (this.dt.Rows.Count <= 0)
        //            return;

        //        this.bindingSource.DataSource = this.dt;

        //        this.dataGridViewTable.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
        //        RowCounter = 0;
        //        foreach (DataRow dr in dt.Rows)
        //        {
        //            //FormBAse = DatabaseQueries.CheckForNull<dynamic>(dr["Base"]);
        //            MessageFromDatabase.Add(dt.Rows.ToString());
        //            for (int ColumnCounter = 0; ColumnCounter < Columnnames.Count; ColumnCounter++)
        //                try
        //                {
        //                    //ArrayMessageFromDatabase[RowCounter, ColumnCounter] = Columnnames[ColumnCounter] + "_" + Utilities.CheckForNullString(DatabaseQueries.CheckForNull<string>(dr[Columnnames[ColumnCounter]]));// = Utilities.CheckForNullString(DatabaseQueries.CheckForNull<string>(dr[Columnnames[counter]]));
        //                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = string.Empty + DatabaseQueries.CheckForNull<dynamic>(dr[Columnnames[ColumnCounter]]);// = Utilities.CheckForNullString(DatabaseQueries.CheckForNull<string>(dr[Columnnames[counter]]));

        //                }//this is helping us to debug and see how the columns are coming out of the database and what data we need.
        //                catch
        //                {
        //                    if (Columnnames[ColumnCounter].Contains("Company"))
        //                        ArrayMessageFromDatabase[RowCounter, ColumnCounter] = CompanyName;
        //                    else if (Columnnames[ColumnCounter].Contains("PO"))
        //                        ArrayMessageFromDatabase[RowCounter, ColumnCounter] = textBox_CustomerPO.Text;
        //                    else if (Columnnames[ColumnCounter].Contains("Form"))
        //                    {
        //                        dynamic TempForm = DatabaseQueries.CheckForNull<dynamic>(dr["Form"]);
        //                        dynamic TempBase = DatabaseQueries.CheckForNull<dynamic>(dr["Base"]);
        //                        dynamic TempCombo = TempForm + TempBase;
        //                        ArrayMessageFromDatabase[RowCounter, ColumnCounter] = TempCombo;
        //                    }
        //                }
        //            RowCounter++;
        //        }
        //    }

        //    catch (Exception e)
        //    {
        //        MessageBox.Show(
        //            "Program Exception: " + e.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        return;
        //    }
        //}
        #endregion commented code

    }
}

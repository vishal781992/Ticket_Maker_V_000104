using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadingApp
{
    class DeclarationClass
    {
        #region Main File Form
        public string Version = "V.0.01.03";//latest as of 04/March/2020
        public string VersionDetails = "Array Size 4000, Cleaned the XML function for tab6, clipboard action added. Added xml support for dominion.\r\nAdded xlsx format for all files .Added Dominion. Cleared all small bugs,\r\nNewest version, Added suport for Kevin everywhere.\r\nAdded more Functions to view the Database.";
      
        public string databaseType = string.Empty;//"dbo";

        public  string ORIGINALSHIPMENTPATH_             = @"\\netserver3\DATA\_ShipmentFiles\";
        public string ROOTDIRFORXMLFILES_               = @"\\Netserver3\DATA\ShipmentsXMLfiles\";
        public string PARENTFOLDERTOSTICKTO_             = @"\\Netserver3\DATA\ShipmentsXMLfiles\";
        public static string SHIPMENTPATH_               = @"\\netserver3\DATA\_ShipmentFiles\";
        public static string VISHALSHIPMENTPATH_         = @"\\netserver3\DATA\Vishal_ShipmentFiles\";
        public static string ROOTDIRFORXMLFILES         = @"\\Netserver3\DATA\ShipmentsXMLfiles"; // the path is used for taking the files
        public static string TICKETLOGDIRECTORY = @"\\netserver3\Data\Log_Tickets_all\TicketLog";

        public static string DividerString = "\r\n-------------------------------------------------";
        public static string DividerString1 = "\r\n-------";
        public static string NotificationString1 = "The File is created! The process is Completed.\r\nSuccess, data Logging\r\nVerify all the columns before sending over to the email.";
        public static string NotificationString2 = "The dataBase might have encountered an error." +
                        "Verification needs data." +
                        "\r\nTry again with correct ticket numbers! ";
        public static string NotificationString3 = "Company has been Successfully Created\r\nYou can process the Pick Ticket For it.";

        public static string ErrorString1 = "The Process is stopped!" +
                        "\r\nEnter The Data mandatory for the Process to continue." +
                        "\r\nSee Ticket for more data" +
                        "\r\nDataBase Name is usually written in handwriting." +
                        "\r\nSold To in the left top of the Ticket includes Company Name" +
                        "\r\nNever use Ship To names." +
                        "\r\nSales Person Codes are coming." +
                        "\r\nqueries to- vishal@visionmetering.com";

        public string FilePathOfXML,
              strFilename1,
              ExportXlSXPath,
              XMLMakerPath,
              Log_DataCollectionString,
              Log_TicketToLog,
              Log_TicketCounts, month_T,
              month_Tminus1,
              month_Tplus1,
              YearForSearch,
              String_SearchDataTab4,
              folderNameForOutputFile,
              FileNameExtension_Global,
              Search_TicketNumber;

        public bool Flag_searchDirectory = false,
            flag_POnumber = true,
            Flag_forDisplayOfDatabase = false,
            IamPopUP = false,
            Flag_searchDirectoryBecauseFoldersUpdated = false,
            Flag_InsertNFHere = true;
        #endregion Main File Form

        public string File1FullPath = string.Empty,
                    File2FullPath = string.Empty,
                    File1Name = string.Empty,
                    File2Name = string.Empty,
                    File1NameTrimmed = string.Empty;

        public static double LowerL_SL_SF = 99.85, HigherL_SL_SF = 100.15, LowerL_SP_WA = 99.75, HigherL_SP_WA = 100.25;

        public string[] CommIDList = new string[] { "CommID", "CommID1", "CommID2", "CommID3", "CommID4" };
        public string[] CommIDShortcuts = new string[] { "C0", "C1", "C2", "C3", "C4" };

    }
}

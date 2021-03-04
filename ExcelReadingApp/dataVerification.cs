using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReadingApp
{
    class dataVerification
    {
        #region Declarations
        DeclarationClass DC = new DeclarationClass();

        protected dynamic[,] arrayMessageFromDatabase = new string[4000, 200];
        public dynamic[,] TempAryForIntegers = new dynamic[4000,25];

        protected List<string> columnValue = new List<string>();
        protected List<string> List_Common_temp = new List<string>();
        protected List<string> List_Common_temp_New = new List<string>();
        protected List<string> meterTypeCodes = new List<string>();

        protected List<string> MeterClassification = new List<string>();

        protected List<double> List_Common_temp_Double = new List<double>();

        protected int rowCounter;
        public int rowcounterTillNow =0, RowCounterForTheSpecTicket, CounterUniversalForEachTicket=0;
        public bool Flag_ErrorInPallet = false;

        //Declaration for AL Values for verification
        
        #endregion Declarations

        #region Constructor important
        public dataVerification(dynamic[,] ArrayMessageFromDatabase, List<string> ColumnValue, int RowCounter,List<string>MeterTypeCodes)
        {
            this.arrayMessageFromDatabase = ArrayMessageFromDatabase;
            this.columnValue = ColumnValue;
            this.rowCounter = RowCounter;
            this.meterTypeCodes = MeterTypeCodes;
        }
        #endregion Constructor important

        #region Verification functions
        public string Verification_ItemRange(string ToFind,string TicketNumber)
        {
            List_Common_temp.Clear(); Flag_ErrorInPallet = false;
            string value = string.Empty;
            int counterToFind = 0, counterToBatch = 0, counterToBox = 0; this.RowCounterForTheSpecTicket = 0;

            try
            {
                while (value != ToFind)
                {
                    counterToFind++;
                    value = columnValue[counterToFind];
                }
                value = string.Empty;
                while (value != "Batch")
                {
                    counterToBatch++;
                    value = columnValue[counterToBatch];
                }
                while (value.ToUpper() != "BOX")
                {
                    counterToBox++;
                    value = columnValue[counterToBox];
                }

                for (int counterD = 0;counterD< rowCounter;counterD++)
                {
                    try
                    {
                        if (arrayMessageFromDatabase[counterD, counterToBatch] == TicketNumber && !string.IsNullOrEmpty(arrayMessageFromDatabase[counterD, counterToBox]))
                        {
                            List_Common_temp.Add(arrayMessageFromDatabase[counterD, counterToFind]);
                            this.RowCounterForTheSpecTicket++;
                        }
                        else
                            Flag_ErrorInPallet = true;
                    }
                    catch { }
                }

                List_Common_temp.Sort();
                bool isParsed1;
                int insequence = 0, notSequence = 0 ;
                if (isParsed1 = int.TryParse(List_Common_temp[0].Substring(List_Common_temp[0].Length-4), out int referenceITEM))
                {
                    foreach (string item in List_Common_temp)
                    {
                        bool isParsed2;
                        if (isParsed2 = int.TryParse(item.Substring(item.Length - 4), out int tempItem))
                        {
                            if (referenceITEM + 1 == tempItem)
                            {
                                insequence++; referenceITEM += 1;
                            }
                            else if (referenceITEM == tempItem) { insequence++; }
                            else { notSequence++; }
                        }
                    }
                }
                if(notSequence>0)
                    return string.Empty + TicketNumber + ":: " + List_Common_temp[0] + " To " + List_Common_temp[List_Common_temp.Count - 1] + ", Not in Sequence";
                else
                    return string.Empty + TicketNumber + ":: " + List_Common_temp[0] + " To " + List_Common_temp[List_Common_temp.Count - 1] + ", In Sequence";
            }
            catch
            {
                MessageBox.Show(TicketNumber +":: "+ ToFind + " Problem with Verification_ItemRange function.");
                return string.Empty;
            }
        }

        public List<string> MeterTypeClassification(List<string>MeterTypeCodes)
        {
            string MeterCode_6 = string.Empty; MeterTypeCodeClassifier MTC = new MeterTypeCodeClassifier();
            foreach (string MeterCode in MeterTypeCodes)
            {
                MeterClassification.Add(MTC.MeterTypeCode001(MeterCode));
            }
            return MeterClassification;
        }

        public Tuple<List<string>,List<double>> Verification_AL_Checks(string ToFind)
        {
            string tempValue = string.Empty; string value = string.Empty;
            int counter = 0, counterToBatch = 0;
            List_Common_temp_Double.Clear(); List_Common_temp.Clear();List_Common_temp_New.Clear();

            while (value.ToUpper() != ToFind.ToUpper())//"FirmwareRevision"
            {
                counter++;
                value = columnValue[counter];
            }
            value = string.Empty;
            while (value.ToUpper() != "METERID")//Batch
            {
                counterToBatch++;
                value = columnValue[counterToBatch];
            }

            if(string.Equals(ToFind,"ALSP") || string.Equals(ToFind, "ALWA"))
            {
                for (int counterD = 0; counterD < rowCounter; counterD++)
                {
                    try
                    {
                        bool result = double.TryParse(arrayMessageFromDatabase[counterD, counter], out double TempDouble);
                        if (result)
                        {
                            if (TempDouble >= DeclarationClass.LowerL_SP_WA && TempDouble <= DeclarationClass.HigherL_SP_WA) { }   //
                            else
                            {
                                List_Common_temp.Add("["+arrayMessageFromDatabase[counterD, counterToBatch] + ": " + arrayMessageFromDatabase[counterD, counter]+"]");
                            }
                        }
                        List_Common_temp_Double.Add(TempDouble);
                    }
                    catch { }
                }
            }

            if (string.Equals(ToFind, "ALSF") || string.Equals(ToFind, "ALSL"))
            {
                for (int counterD = 0; counterD < rowCounter; counterD++)
                {
                    try
                    {
                        bool result = double.TryParse(arrayMessageFromDatabase[counterD, counter], out double TempDouble);
                        if (result)
                        {
                            if (TempDouble >= DeclarationClass.LowerL_SL_SF && TempDouble <= DeclarationClass.HigherL_SL_SF) { }
                            else
                            {
                                List_Common_temp.Add("["+arrayMessageFromDatabase[counterD, counterToBatch] + ": " + arrayMessageFromDatabase[counterD, counter]+"]");
                            }
                        }
                        List_Common_temp_Double.Add(TempDouble);
                    }
                    catch { }
                }
            }
            //sorting the Lists here.
            List_Common_temp.Sort(); List_Common_temp_Double.Sort();

            return new Tuple<List<string>, List<double>>(List_Common_temp, List_Common_temp_Double);
        }

        public List<string> VerificationOfCommID()
        {
            Array.Clear(TempAryForIntegers, 0, TempAryForIntegers.Length);
            List_Common_temp_Double.Clear(); List_Common_temp.Clear(); List_Common_temp_New.Clear();
            string tempValue = string.Empty; string value = string.Empty;
            int counter = 0, counterToBatch = 0;
            try
            {
                while (value.ToUpper() != "BATCH")//Batch
                {
                    counterToBatch++;
                    value = columnValue[counterToBatch];

                }
                for (int LoopCounter = 0;LoopCounter<DC.CommIDList.Length;LoopCounter++)
                {
                    counter = 0;
                    try
                    {
                        while (value.ToUpper() != DC.CommIDList[LoopCounter].ToUpper() || counter == columnValue.Count)//"FirmwareRevision"
                        {
                            counter++;
                            value = columnValue[counter];
                        }
                        if (counter != columnValue.Count && !string.IsNullOrEmpty(arrayMessageFromDatabase[1, counter]))
                        {
                            for (int counterD = 0; counterD < rowCounter; counterD++)
                            {
                                try
                                {
                                    string TempString = arrayMessageFromDatabase[counterD, counter];
                                    if (TempString.StartsWith("05") || TempString.StartsWith("08")) { }
                                    else 
                                    {
                                        if(string.IsNullOrEmpty(List_Common_temp[0]) || !List_Common_temp[0].StartsWith(TempString.Substring(0, 2)))
                                            List_Common_temp.Add("Starts with: "+TempString.Substring(0,2)); 
                                    }//List_Common_temp.Add("[" + arrayMessageFromDatabase[counterD, counterToBatch] + "(" + CommIDShortcuts[LoopCounter] + "): " + TempString + "]");
                                }
                                catch { }
                            }
                        }
                    }
                    catch
                    {}
                }
                return List_Common_temp;
            }
            catch { return List_Common_temp; }
        }

        public string DuplicateRecordVerification(string ToFind,string Database, string dbo_type,[Optional]string WhereCondnString)
        {
            Array.Clear(TempAryForIntegers, 0, TempAryForIntegers.Length);
            List_Common_temp.Clear(); List_Common_temp_New.Clear();
            string tempValue = string.Empty; string value = string.Empty; string ReturnStringWithErrors = string.Empty; string TempStringForConCat = string.Empty;
            int counter = 0, counterToMeterID = 0, counterD = 0;//cmid

            QueryTest QT = new QueryTest();
            string TempConnectionString = "Server=" + "Netserver3" + "; Database=" + Database + "; UId=" + "power" + "; Password=" + "power" + ";";//master
            if (string.IsNullOrEmpty(WhereCondnString))
                WhereCondnString = "METERID";
            while (value.ToUpper() != WhereCondnString.ToUpper())
            {
                counterToMeterID++;
                value = columnValue[counterToMeterID];
            }

            try
            {
                while (counterD < rowCounter)
                {
                    List_Common_temp_New = QT.DuplicateCheckInDB(ToFind, arrayMessageFromDatabase[counterD, counterToMeterID], columnValue[counterToMeterID], TempConnectionString, dbo_type);
                    if(QT.Flag_DuplicateRecord)
                    {
                        foreach (string rec in List_Common_temp_New)
                            TempStringForConCat += "-"+rec;
                        List_Common_temp.Add(TempStringForConCat);
                    }
                   
                    List_Common_temp_New.Clear(); TempStringForConCat = string.Empty;//cleaar the cache
                    counterD++;
                }
                counter++;
                
                foreach (string str in List_Common_temp)
                {
                    if(!string.IsNullOrEmpty(str))
                        ReturnStringWithErrors += "-" + str;
                }

                if (string.IsNullOrEmpty(ReturnStringWithErrors))
                    ReturnStringWithErrors = " No Dup.";

                return ReturnStringWithErrors;
            }
            catch{ ReturnStringWithErrors = "Catch occured, Check with Programmer!"; return ReturnStringWithErrors; }
        }

        public List<string> Verification_General_typeSort(string ToFind, string TicketNumber)
        {
            int TempAryForIntegers_element_A = 0; Array.Clear(TempAryForIntegers, 0, TempAryForIntegers.Length); int tempVar = 0;
            List_Common_temp.Clear();string tempValue = string.Empty;string value = string.Empty;
            int counter = 0, counterToBatch = 0;
            try
            {
                
                while (value.ToUpper() != ToFind.ToUpper())//"FirmwareRevision"
                {
                    counter++;
                    value = columnValue[counter];
                }
                value = string.Empty;
                while (value.ToUpper() != "BATCH")//Batch
                {
                    counterToBatch++;
                    value = columnValue[counterToBatch];

                }
                
                for (int counterD = 0; counterD < rowCounter; counterD++)
                {
                    try
                    {
                        if (arrayMessageFromDatabase[counterD, counterToBatch] == TicketNumber)
                        {
                            if (counterD == 0 || !string.Equals(arrayMessageFromDatabase[counterD, counter].ToUpper(), tempValue.ToUpper()))
                            {
                                if (arrayMessageFromDatabase[counterD, counterToBatch] == TicketNumber)
                                {
                                    tempValue = arrayMessageFromDatabase[counterD, counter];
                                    List_Common_temp.Add(arrayMessageFromDatabase[counterD, counter]);
                                    TempAryForIntegers_element_A++; tempVar = 0;
                                }
                            }
                            tempVar++;
                            TempAryForIntegers[TempAryForIntegers_element_A, 0] = tempVar;
                        }
                    }
                    catch { }
                }

                #region commented
                //tempRC = RowCounterForEachbatchFile(counterToBatch, TicketNumber);

                //for (int loopCounter = rowcounterTillNow; loopCounter < rowcounterTillNow + tempRC; loopCounter++)//rowcounterTillNow + 1
                //    {
                //    if (loopCounter == rowcounterTillNow || !string.Equals(arrayMessageFromDatabase[loopCounter, counter].ToUpper(), tempValue.ToUpper()))
                //    {
                //        if (arrayMessageFromDatabase[loopCounter, counterToBatch] == TicketNumber)
                //        {
                //            tempValue = arrayMessageFromDatabase[loopCounter, counter];
                //            List_Common_temp.Add(arrayMessageFromDatabase[loopCounter, counter]);
                //        }
                //    }
                //}
                #endregion commented
                int counterForeachLoop = 1;
                foreach (string item in List_Common_temp)
                {
                    List_Common_temp_New.Add(item + "(" + TempAryForIntegers[counterForeachLoop, 0] + ")");
                        counterForeachLoop++;
                }
                //List_Common_temp.Sort();//temp commented
                return List_Common_temp_New;
            }
            catch
            {
                MessageBox.Show(TicketNumber+ ":: " + ToFind + " Problem with Verification_General_typeSort function.");
                return List_Common_temp_New;
            }
        }

        public int RowCounterForEachbatchFile(int counterToBatch, string TicketNumber)
        {
            RowCounterForTheSpecTicket = 0;
            
            for (int loopCounter = 0; loopCounter < rowCounter; loopCounter++)
            {
                if (arrayMessageFromDatabase[loopCounter, counterToBatch] == TicketNumber)
                    RowCounterForTheSpecTicket += 1;
            }
            return RowCounterForTheSpecTicket;
        }

        public void CounterGenerator() { this.CounterUniversalForEachTicket++; }
        #endregion Verification functions
    }
}

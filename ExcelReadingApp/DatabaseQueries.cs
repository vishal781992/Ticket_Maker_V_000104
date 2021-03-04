using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelReadingApp
{
    #region SQLManager Class

    public static class DatabaseQueries
    {
        #region Variables

        public static bool SQLExpressProblem = false;

        public static bool MeterTestHistoryTable = false;
        public static bool MeterTestCustomTable = false;
        public static bool MeterKVARTestTable = false;
        public static bool UserLogTable = false;
        public static int MaximumCustomFields = 0;
        public static bool PCBlot = false;
        public static bool VestaVersion = false;
        public static bool VestaFirmware = false;
        public static bool VestaScript = false;

        public static bool AFKW = false;                    // FlorenceDB
        public static bool ALKW = false;                    // FlorenceDB
        public static bool AFKWH = false;                   // OzarksDB / FlorenceDB
        public static bool ALKWH = false;                   // OzarksDB / FlorenceDB
        public static bool ModuleKWH = false;               // OzarksDB / FlorenceDB
        public static bool CompanyNo = false;               // FlorenceDB

        public static bool AFKWMeterTest = false;           // CarrollDB
        public static bool ALKWMeterTest = false;           // CarrollDB
        public static bool AFKWHMeterTest = false;          // CarrollDB
        public static bool ALKWHMeterTest = false;          // CarrollDB
        public static bool AFKVARMeterTest = false;         // CarrollDB
        public static bool ALKVARMeterTest = false;         // CarrollDB
        public static bool StatusCodeMeterTest = false;     // CarrollDB

        public static bool VestaDatabase = false;
        public static bool StatisticsTable = false;
        public static string VestaStatisticsTableName = "Stats";

        //public static string VestaConnectionString = @"Server=.\SQLExpress; Database=VestaDB; Integrated Security=yes; Uid=auth_windows;";
        public static string VestaConnectionString = @"Server=.\SQLExpress; Database=VestaDB; UId=sa; Password=visionmetering;";

        #endregion Variables

        #region Check For Null

        public static T CheckForNull<T>(object obj)
        {
            if (obj == null || obj == DBNull.Value)
            {
                return default(T); // returns the default value for the type
            }
            else
            {
                return (T)obj;
            }
        }

        public static string CheckForNull(object obj)
        {
            if (!DBNull.Value.Equals(obj))
                return (string)obj;
            else return string.Empty;
        }

        public static long CheckForNullForInt(object obj)
        {
            if (!DBNull.Value.Equals(obj))
                return (long)obj;
            else return 0;
        }

        #endregion Check For Null

        #region Load Parameter

        public static void LoadParameter(SqlCommand command, string field, SqlDbType type, object value)
        {
            command.Parameters.Add(field, type);
            command.Parameters[field].Value = value;

            //VestaDLL.DebugManager.UpdateLog("SQLManager.LoadParameter", "field=" + field + " value=" + value.ToString());
        }

        #endregion Load Parameter

        #region Execute Query



        //public static DataTable ExecuteQuery(string query)
        //{
        //    DataTable table = new DataTable();

        //    if (SQLManager.User.ConnectionString == string.Empty)
        //        SQLManager.User.SetConnectionString(SQLManager.User.Server, SQLManager.User.Database, SQLManager.User.DBCredentials.UserID, SQLManager.User.DBCredentials.Password);

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(SQLManager.User.ConnectionString))
        //        {
        //            using (SqlCommand cmd = new SqlCommand(query, conn))
        //            {
        //                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
        //                conn.Open();

        //                adapter.Fill(table);

        //                conn.Close();
        //            }
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }

        //    return table;
        //}

        public static DataTable ExecuteQuery(string query, string connectionString)
        {
            DataTable table = new DataTable();

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                        conn.Open();

                        adapter.Fill(table);

                        conn.Close();
                    }
                }
            }

            catch// (Exception ex)
            {
                //MessageBox.Show(
                //    ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
                //    "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return table;

        }

        #endregion Execute Query

        #region Get Databases
        public static DataTable GetDatabases(string connectionString)
        {
            DataTable dt;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                // Get the schema information of Databases in your instance
                dt = conn.GetSchema("Databases");
            }

            return dt;
        }

        #endregion Get Databases

        #region Find Table

        //private static bool FindTable(string tableName)
        //{
        //    //+VestaDLL.DebugManager.UpdateLog("Enter FindTable");

        //    int count = 0;

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(StatusInformation.User.ConnectionString))
        //        {
        //            conn.Open();

        //            string[] columnRestrictions = new string[4];

        //            // For the array,
        //            //    0-member represents Catalog;
        //            //    1-member represents Schema;
        //            //    2-member represents Table Name;
        //            //    3-member represents Column Name.

        //            // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
        //            columnRestrictions[2] = tableName;

        //            DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

        //            var selectedRows =
        //                from info in schemaTable.AsEnumerable()
        //                select new
        //                {
        //                    TableCatalog = info["TABLE_CATALOG"],
        //                    TableSchema = info["TABLE_SCHEMA"],
        //                    TableName = info["TABLE_NAME"],
        //                    ColumnName = info["COLUMN_NAME"],
        //                    DataType = info["DATA_TYPE"]
        //                };

        //            foreach (var row in selectedRows)
        //                count++;

        //            conn.Close();
        //        }

        //        //+VestaDLL.DebugManager.UpdateLog("FindTable", "count=" + count);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindTable");

        //        if (count > 0)
        //            return true;
        //        else
        //            return false;
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("FindMeterTestHistoryTable", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindMeterTestHistoryTable");

        //        return false;
        //    }
        //}

        private static bool FindTable(string tableName, string connectionString)
        {
            //+VestaDLL.DebugManager.UpdateLog("Enter FindTable");

            int count = 0;

            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    string[] columnRestrictions = new string[4];

                    // For the array,
                    //    0-member represents Catalog;
                    //    1-member represents Schema;
                    //    2-member represents Table Name;
                    //    3-member represents Column Name.

                    // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
                    columnRestrictions[2] = tableName;

                    DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

                    var selectedRows =
                        from info in schemaTable.AsEnumerable()
                        select new
                        {
                            TableCatalog = info["TABLE_CATALOG"],
                            TableSchema = info["TABLE_SCHEMA"],
                            TableName = info["TABLE_NAME"],
                            ColumnName = info["COLUMN_NAME"],
                            DataType = info["DATA_TYPE"]
                        };

                    foreach (var row in selectedRows)
                        count++;

                    conn.Close();
                }

                //+VestaDLL.DebugManager.UpdateLog("FindTable", "count=" + count);
                //+VestaDLL.DebugManager.UpdateLog("Leave FindTable");

                if (count > 0)
                    return true;
                else
                    return false;
            }

            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
                    "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //+VestaDLL.DebugManager.UpdateLog("FindMeterTestHistoryTable", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
                //+VestaDLL.DebugManager.UpdateLog("Leave FindMeterTestHistoryTable");

                return false;
            }
        }

        #endregion Find Table

        #region Find Field

        //private static bool FindField(string tableName, string fieldName)
        //{
        //    //+VestaDLL.DebugManager.UpdateLog("Enter FindField");

        //    //int count = 0;

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(StatusInformation.User.ConnectionString))
        //        {
        //            conn.Open();

        //            string[] columnRestrictions = new string[4];

        //            // For the array,
        //            //    0-member represents Catalog;
        //            //    1-member represents Schema;
        //            //    2-member represents Table Name;
        //            //    3-member represents Column Name.

        //            // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
        //            columnRestrictions[2] = tableName;

        //            DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

        //            var selectedRows =
        //                from info in schemaTable.AsEnumerable()
        //                select new
        //                {
        //                    TableCatalog = info["TABLE_CATALOG"],
        //                    TableSchema = info["TABLE_SCHEMA"],
        //                    TableName = info["TABLE_NAME"],
        //                    ColumnName = info["COLUMN_NAME"],
        //                    DataType = info["DATA_TYPE"]
        //                };

        //            foreach (var row in selectedRows)
        //            {
        //                string columnName = row.ColumnName.ToString();

        //                if (columnName == fieldName)
        //                {
        //                    //+VestaDLL.DebugManager.UpdateLog("FindField", "'PCBlot' found");
        //                    //+VestaDLL.DebugManager.UpdateLog("Leave FindField");
        //                    return true;
        //                }
        //            }

        //            conn.Close();
        //        }

        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindField");
        //        return false;
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("FindField", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindField");

        //        return false;
        //    }
        //}

        //private static bool FindTable(string tableName)
        //{
        //    //+VestaDLL.DebugManager.UpdateLog("Enter FindTable");

        //    int count = 0;

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(SQLManager.User.ConnectionString))
        //        {
        //            conn.Open();

        //            string[] columnRestrictions = new string[4];

        //            // For the array,
        //            //    0-member represents Catalog;
        //            //    1-member represents Schema;
        //            //    2-member represents Table Name;
        //            //    3-member represents Column Name.

        //            // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
        //            columnRestrictions[2] = tableName;

        //            DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

        //            var selectedRows =
        //                from info in schemaTable.AsEnumerable()
        //                select new
        //                {
        //                    TableCatalog = info["TABLE_CATALOG"],
        //                    TableSchema = info["TABLE_SCHEMA"],
        //                    TableName = info["TABLE_NAME"],
        //                    ColumnName = info["COLUMN_NAME"],
        //                    DataType = info["DATA_TYPE"]
        //                };

        //            foreach (var row in selectedRows)
        //                count++;

        //            conn.Close();
        //        }

        //        //+VestaDLL.DebugManager.UpdateLog("FindTable", "count=" + count);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindTable");

        //        if (count > 0)
        //            return true;
        //        else
        //            return false;
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("FindMeterTestHistoryTable", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindMeterTestHistoryTable");

        //        return false;
        //    }
        //}

        //private static bool FindTable(string tableName, string connectionString)
        //{
        //    //+VestaDLL.DebugManager.UpdateLog("Enter FindTable");

        //    int count = 0;

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(connectionString))
        //        {
        //            conn.Open();

        //            string[] columnRestrictions = new string[4];

        //            // For the array,
        //            //    0-member represents Catalog;
        //            //    1-member represents Schema;
        //            //    2-member represents Table Name;
        //            //    3-member represents Column Name.

        //            // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
        //            columnRestrictions[2] = tableName;

        //            DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

        //            var selectedRows =
        //                from info in schemaTable.AsEnumerable()
        //                select new
        //                {
        //                    TableCatalog = info["TABLE_CATALOG"],
        //                    TableSchema = info["TABLE_SCHEMA"],
        //                    TableName = info["TABLE_NAME"],
        //                    ColumnName = info["COLUMN_NAME"],
        //                    DataType = info["DATA_TYPE"]
        //                };

        //            foreach (var row in selectedRows)
        //                count++;

        //            conn.Close();
        //        }

        //        //+VestaDLL.DebugManager.UpdateLog("FindTable", "count=" + count);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindTable");

        //        if (count > 0)
        //            return true;
        //        else
        //            return false;
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("FindMeterTestHistoryTable", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindMeterTestHistoryTable");

        //        return false;
        //    }
        //}

        //#endregion Find Table

        //#region Find Field

        //private static bool FindField(string tableName, string fieldName)
        //{
        //    //+VestaDLL.DebugManager.UpdateLog("Enter FindField");

        //    //int count = 0;

        //    try
        //    {
        //        using (SqlConnection conn = new SqlConnection(SQLManager.User.ConnectionString))
        //        {
        //            conn.Open();

        //            string[] columnRestrictions = new string[4];

        //            // For the array,
        //            //    0-member represents Catalog;
        //            //    1-member represents Schema;
        //            //    2-member represents Table Name;
        //            //    3-member represents Column Name.

        //            // Now we specify the Table_Name and Column_Name of the columns what we want to get schema information.
        //            columnRestrictions[2] = tableName;

        //            DataTable schemaTable = conn.GetSchema("Columns", columnRestrictions);

        //            var selectedRows =
        //                from info in schemaTable.AsEnumerable()
        //                select new
        //                {
        //                    TableCatalog = info["TABLE_CATALOG"],
        //                    TableSchema = info["TABLE_SCHEMA"],
        //                    TableName = info["TABLE_NAME"],
        //                    ColumnName = info["COLUMN_NAME"],
        //                    DataType = info["DATA_TYPE"]
        //                };

        //            foreach (var row in selectedRows)
        //            {
        //                string columnName = row.ColumnName.ToString();

        //                if (columnName == fieldName)
        //                {
        //                    //+VestaDLL.DebugManager.UpdateLog("FindField", "'PCBlot' found");
        //                    //+VestaDLL.DebugManager.UpdateLog("Leave FindField");
        //                    return true;
        //                }
        //            }

        //            conn.Close();
        //        }

        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindField");
        //        return false;
        //    }

        //    catch (Exception ex)
        //    {
        //        Utilities.ShowMessageBox(
        //            ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("FindField", ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace);
        //        //+VestaDLL.DebugManager.UpdateLog("Leave FindField");

        //        return false;
        //    }
        //}

        #endregion Find Field

    }

    public class QueryTest
    {
        #region init
        private User user;
        private DataTable dt;
        private BindingSource bindingSource = new BindingSource();

        public List<string> MeterTypeCodes = new List<string>();
        public dynamic[,] ArrayMessageFromDatabase = new dynamic[4000,200]; string MessageFromDatabase;
        public List<string> MessageFromDatabaseList = new List<string>();
        public bool Flag_DuplicateRecord;

        string[] AryOfColumns = new string[4000];

        public int RowCounter { get; set; }
        #endregion init

        #region User init
        public void USER_init(string comboBox_DataBaseName)
        {
            this.user = new User();
            user.Server = "Netserver3";
            user.Database = comboBox_DataBaseName; 
            user.DBOwner = "dbo";
            user.SQLCredentials = new Credentials();
            user.SQLCredentials.UserID = "power";
            user.SQLCredentials.Password = "power";

            user.SetConnectionString();  //connection string is set here
        }
        #endregion User init

        #region GetDataTables
        public DataTable GetDataTables()
        {
            DataTable dt;
            string ConnectionString = "Server=" + "Netserver3" + "; Database=" + "master" + "; UId=" + "power" + "; Password=" + "power" + ";";
            dt = DatabaseQueries.GetDatabases(ConnectionString);

            return dt;
        }
        #endregion GetDataTables

        #region Tab5_AllDataQuery
        public DataTable Tab5_AllDataQuery(string TicketNumberString , string Database, string dbo_type)
        {
            try
            {
                string Server = "netserver3";

                string tempA = "SELECT * " +
                    "FROM ((" + dbo_type + ".Meter INNER JOIN " + dbo_type + ".MeterTypeView ON " + dbo_type + ".Meter.MeterTypeCode = " + dbo_type + ".MeterTypeView.MeterTypeCode) " +
                    "INNER JOIN " + dbo_type + ".MeterTest ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterTest.MeterID) " +
                    "INNER JOIN " + dbo_type + ".MeterReadings ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterReadings.MeterID " +
                    "WHERE (((" + dbo_type + ".Meter.Batch)='" + TicketNumberString + "')) ";
                string tempC = "ORDER BY " + dbo_type + ".Meter.MeterID, " + dbo_type + ".Meter.Box, " + dbo_type + ".Meter.Pallet," + dbo_type + ".Meter.IMEI";

                string query = tempA + tempC;

                string ConnectionString = "Server=" + Server + "; Database=" + Database + "; UId=" + "power" + "; Password=" + "power" + ";";

                this.dt = DatabaseQueries.ExecuteQuery(query, ConnectionString);
                if (this.dt.Rows.Count <= 0)
                    return dt;

                dynamic[] temp3 = new dynamic[500];

                for (int count1 = 0; count1 < dt.Columns.Count; count1++)
                {
                    bool temp1 = dt.Rows[count1].IsNull(count1);
                    if (temp1)
                    {
                        dt.Columns.RemoveAt(count1);
                        count1 -= 1;
                    }
                    dynamic temp2 = dt.Rows[count1 + 1];
                    dt.Rows.CopyTo(temp3, 0);
                }

                dt.AcceptChanges();

                return dt;
            }
            catch 
            {
                return dt; 
            }
        }
        #endregion Tab5_AllDataQuery

        #region Tab5_ColumnNameQuery
        public string[] Tab5_ColumnNameQuery(string TicketNumberString, string Database, string ToFind ,string dbo_type)
        {
            string Server = "netserver3";

            string tempA = "SELECT * " +
                "FROM ((" + dbo_type + ".Meter INNER JOIN " + dbo_type + ".MeterTypeView ON " + dbo_type + ".Meter.MeterTypeCode = " + dbo_type + ".MeterTypeView.MeterTypeCode) " +
                "INNER JOIN " + dbo_type + ".MeterTest ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterTest.MeterID) " +
                "INNER JOIN " + dbo_type + ".MeterReadings ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterReadings.MeterID " +
                "WHERE (((" + dbo_type + ".Meter.Batch)='" + TicketNumberString + "')) ";
            string tempC = "ORDER BY " + dbo_type + ".Meter.MeterID, " + dbo_type + ".Meter.Box, " + dbo_type + ".Meter.Pallet," + dbo_type + ".Meter.IMEI";

            string query = tempA + tempC;


            string ConnectionString =
                "Server=" + Server + "; Database=" + Database + "; UId=" + "power" + "; Password=" + "power" + ";";

            dt = DatabaseQueries.ExecuteQuery(query, ConnectionString);
            int count = 0;
            foreach (dynamic d in dt.Columns)
            {
                dynamic ColumnHead = d;
                try
                {
                    AryOfColumns[count] = ColumnHead + string.Empty;count++;
                }
                catch { }
            }
            return AryOfColumns;
        }
        #endregion Tab5_ColumnNameQuery

        #region Tab1_TestQuery
        public void Tab1_TestQuery(List<string> Columnnames, string CompanyName,List<string> TicketsList, string textBox_CustomerPO,string dbo_type)
        {
            #region init
            string TicketNumberString = string.Empty;
            string PennyAdder = "or";
            if (TicketsList.Count>1)
            {
                for(int counter =0;counter< TicketsList.Count;counter++)
                {
                    if (counter < TicketsList.Count - 1)
                        TicketNumberString = TicketNumberString + TicketsList[counter] + PennyAdder;
                    if (counter == TicketsList.Count - 1)
                        TicketNumberString = TicketNumberString + TicketsList[counter];
                }
                
            }
            else
                TicketNumberString = TicketsList[0];
            #endregion init

            string tempA = "SELECT * " +
                "FROM ((" + dbo_type + ".Meter INNER JOIN " + dbo_type + ".MeterTypeView ON " + dbo_type + ".Meter.MeterTypeCode = " + dbo_type + ".MeterTypeView.MeterTypeCode) " +
                "INNER JOIN " + dbo_type + ".MeterTest ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterTest.MeterID) " +
                "INNER JOIN " + dbo_type + ".MeterReadings ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterReadings.MeterID " +
                "WHERE ";

            string tempC = string.Empty;

            string query = tempA + TicketNumberString;
            //last version shares all the comments, beta 1 ...ExcelReadingApp_12_v000100_working_Stable0102 version_beta_1

            try
            {
                this.dt = DatabaseQueries.ExecuteQuery(query, user.ConnectionString);
                if (this.dt.Rows.Count <= 0)
                    return;

                RowCounter = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    MeterTypeCodes.Add(string.Empty + DatabaseQueries.CheckForNull<dynamic>(dr["MeterTypeCode"]));
                    for (int ColumnCounter = 0; ColumnCounter < Columnnames.Count; ColumnCounter++)
                    {
                        try
                        {
                            ArrayMessageFromDatabase[RowCounter, ColumnCounter] = string.Empty + DatabaseQueries.CheckForNull<dynamic>(dr[Columnnames[ColumnCounter]]);
                        } //this is helping us to debug and see how the columns are coming out of the database and what data we need.
                        catch
                        {
                            if (Columnnames[ColumnCounter].Contains("Company"))
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = CompanyName;
                            else if (Columnnames[ColumnCounter].Contains("PO"))
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = textBox_CustomerPO;
                            else if (Columnnames[ColumnCounter].Contains("Form"))
                            {
                                dynamic TempForm = DatabaseQueries.CheckForNull<dynamic>(dr["Form"]);
                                dynamic TempBase = DatabaseQueries.CheckForNull<dynamic>(dr["Base"]);
                                dynamic TempCombo = TempForm + TempBase;    //concatination
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = TempCombo;
                            }
                            else if (Columnnames[ColumnCounter].Contains("8digitCommID"))
                            {
                                dynamic TempCombo = string.Empty + DatabaseQueries.CheckForNull<dynamic>(dr["CommID"]);
                                string demoString = TempCombo.ToString();
                                if(demoString.Length>8)
                                {
                                    int TotalSizeOfCommID = demoString.Length;
                                    int CharsToRemove = TotalSizeOfCommID - 8;
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = demoString.Substring(CharsToRemove);
                                }
                                else
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = demoString;
                            }
                        }
                    }
                    RowCounter++; 
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Test Quest, Program Exception: "+ dt.Rows.Count+"-------" + e.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            dt.Dispose();
        }
        #endregion Tab1_TestQuery

        #region DuplicateCheckInDB
        public List<string> DuplicateCheckInDB(string ToFind, string MeterID, string KeyWord, string TempConnectionString, string dbo_type)
        {
            
            string query = "select "+ dbo_type + ".Meter." + ToFind + " from "+ dbo_type + ".Meter where "+ dbo_type + ".Meter." + KeyWord + " =" + "'" + MeterID + "'"; //Batch, MeterID

            try
            {
                this.dt = DatabaseQueries.ExecuteQuery(query, TempConnectionString);
                if (this.dt.Rows.Count <= 1)
                {
                    int count = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageFromDatabaseList.Add(DatabaseQueries.CheckForNull<string>(dr[ToFind]));count++;//DatabaseQueries.CheckForNull<>
                    }
                    this.Flag_DuplicateRecord = false;//debug only, else false
                    return MessageFromDatabaseList;
                }
                   
                else
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageFromDatabase = DatabaseQueries.CheckForNull<string>(dr[ToFind]); this.Flag_DuplicateRecord = true;
                        return MessageFromDatabaseList;
                    }
                }
            }
            catch { MessageBox.Show("Error in the Duplicate Finder."); }
            return MessageFromDatabaseList;
        }
        #endregion DuplicateCheckInDB

        #region TestQuerySpcl
        public void TestQuerySpcl(List<string> Spcl_DatDBColumnNames, List<string> Spcl_FileColumnNames, List<string> Spcl_ValueForColumnStatics, List<string> Spcl_MergeEvents,string textbox_t6_ticket, string dbo_type, string WhatToFind, string PONumber)
        {
            //string dbo_type = databaseType;
            string tempA = "SELECT * " +
                "FROM (((" + dbo_type + ".Meter INNER JOIN " + dbo_type + ".MeterTypeView ON " + dbo_type + ".Meter.MeterTypeCode = " + dbo_type + ".MeterTypeView.MeterTypeCode) " +
                "INNER JOIN " + dbo_type + ".MeterTest ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterTest.MeterID) " +
                "INNER JOIN " + dbo_type + ".MeterReadings ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterReadings.MeterID) " +
                "WHERE (" + dbo_type+".Meter."+ WhatToFind + ")='" + textbox_t6_ticket + "'";

            /*            string tempA = "SELECT * " +
                "FROM ((" + dbo_type + ".Meter INNER JOIN " + dbo_type + ".MeterTypeView ON " + dbo_type + ".Meter.MeterTypeCode = " + dbo_type + ".MeterTypeView.MeterTypeCode) " +
                "INNER JOIN " + dbo_type + ".MeterTest ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterTest.MeterID) " +
                "INNER JOIN " + dbo_type + ".MeterReadings ON " + dbo_type + ".Meter.MeterID = " + dbo_type + ".MeterReadings.MeterID " +
                "INNER JOIN " + dbo_type + ".TestSetup ON " + dbo_type + ".MeterType.TestSetupCode = " + dbo_type + ".TestSetup.TestSetupCode " +

                "WHERE (" + dbo_type+".Meter."+ WhatToFind + ")='" + textbox_t6_ticket + "'";
             
             */
            string query = tempA;//+ tempC;

            this.dt = DatabaseQueries.ExecuteQuery(query, user.ConnectionString);
            if (this.dt.Rows.Count <= 0)
                return;

            //this.bindingSource.DataSource = this.dt;

            RowCounter = 0;
            foreach (DataRow drElement in dt.Rows)
            {
                try
                {
                    for (int ColumnCounter = 0; ColumnCounter < Spcl_DatDBColumnNames.Count; ColumnCounter++)
                    {
                        if (!string.IsNullOrEmpty(Spcl_DatDBColumnNames[ColumnCounter]) && string.IsNullOrEmpty(Spcl_ValueForColumnStatics[ColumnCounter]) && string.IsNullOrEmpty(Spcl_MergeEvents[ColumnCounter]))//r1,r3N,r4
                        {
                            try
                            {
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                            }
                            catch
                            { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }
                        }
                        else if (string.IsNullOrEmpty(Spcl_DatDBColumnNames[ColumnCounter]) && !string.IsNullOrEmpty(Spcl_ValueForColumnStatics[ColumnCounter]))//r1N,r3
                        {
                            ArrayMessageFromDatabase[RowCounter, ColumnCounter] = string.Empty + Spcl_ValueForColumnStatics[ColumnCounter];
                        }
                        //formatting row is processed here
                        if (!string.IsNullOrEmpty(Spcl_DatDBColumnNames[ColumnCounter]) && !string.IsNullOrEmpty(Spcl_MergeEvents[ColumnCounter]))
                        {
                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("ADD1"))
                            {
                                try
                                {
                                    string temp = Spcl_MergeEvents[ColumnCounter].Substring(Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD1(") + 5, Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf(")") - Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD1(") - 5);
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = temp + DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }
                                
                            }
                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("ADD2"))
                            {
                                try
                                {
                                    string temp = Spcl_MergeEvents[ColumnCounter].Substring(Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD2(") + 5, Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf(")") - Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD2(") - 5);
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]) + temp;
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }
                            }
                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("ADD3"))
                            {
                                try
                                {
                                    string temp = Spcl_MergeEvents[ColumnCounter].Substring(Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD3(") + 5, Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf(")") - Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("ADD3(") - 5);
                                    string[] elements = temp.Split(',');
                                    foreach (string ELE in elements)
                                    {
                                        temp = ELE.Trim(')', '(', ',');
                                        ArrayMessageFromDatabase[RowCounter, ColumnCounter] += string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[temp]);
                                    }
                                }
                                catch { }
                            }


                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("DEL1"))
                            {
                                try
                                {
                                    string temp = Spcl_MergeEvents[ColumnCounter].Substring(Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("DEL1(") + 5, Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf(")") - Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("DEL1(") - 5);
                                    if(int.TryParse(temp, out int resultINT))
                                    {
                                        temp = DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                        temp = temp.Substring(resultINT);   //trimming the lenght mentioned in the request
                                    }
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = temp;
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }

                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("FORMAT"))
                            {
                                string temp01 = string.Empty; double IntermDouble = 0;
                                string temp = Spcl_MergeEvents[ColumnCounter].Substring(Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("FORMAT(") + 7, Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf(")") - Spcl_MergeEvents[ColumnCounter].ToUpper().IndexOf("FORMAT(") - 7);
                                dynamic result = DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                try
                                {
                                    double.TryParse(result, out IntermDouble);
                                }
                                catch { IntermDouble = result; }

                                temp01 = string.Format(temp, IntermDouble);
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = temp01; 
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("DATE"))
                            {
                                string result = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                try
                                {
                                    result = result.Substring(0, result.IndexOf(' '));
                                }
                                catch { result = string.Empty; }
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = result; 
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("CUSTOMDATE"))
                            {
                                string[] tempAry = Spcl_MergeEvents[ColumnCounter].Split('_');
                                string result = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                if(DateTime.TryParse(result, out DateTime result_1))
                                {
                                    try
                                    {
                                        result = result_1.ToString(tempAry[1]);
                                    }
                                    catch { result = string.Empty; }
                                }
                                
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = result;
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("DATE_REV"))
                            {
                                string result = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[Spcl_DatDBColumnNames[ColumnCounter]]);
                                try
                                {
                                    result = result.Substring(0, result.IndexOf(' '));
                                    string[]DateFromDB = result.Split('/');
                                    result = DateFromDB[2] + "/" + DateFromDB[0] + "/" + DateFromDB[1];
                                }
                                catch { result = string.Empty; }
                                ArrayMessageFromDatabase[RowCounter, ColumnCounter] = result;
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("DATAROW"))
                            {
                                try
                                {
                                    int temp = dt.Rows.Count;
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = temp+string.Empty;
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("TAB"))
                            {
                                try
                                {
                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = PONumber;
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = null; }
                            }

                            if (Spcl_MergeEvents[ColumnCounter].ToUpper().Contains("MERGE"))
                            {
                                string[] ColumnNames = Spcl_MergeEvents[ColumnCounter].Split(',');
                                try
                                {
                                    ColumnNames[0] = ColumnNames[0].Substring(6);
                                    ColumnNames[2] = ColumnNames[2].Trim(')');
                                }
                                catch { }

                                try
                                {
                                    string result = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[ColumnNames[0]]);
                                    string result1 = string.Empty + DatabaseQueries.CheckForNull<dynamic>(drElement[ColumnNames[2]]); 
                                    string result2 = string.Empty;

                                    result2 = ColumnNames[1].Contains("NS") ? (result + result1) : (result + ColumnNames[1] + result1);

                                    ArrayMessageFromDatabase[RowCounter, ColumnCounter] = result2;
                                }
                                catch { ArrayMessageFromDatabase[RowCounter, ColumnCounter] = "Error"; }
                            }
                        }
                    }
                }
                catch { }
                RowCounter++; 
            }
        }
        #endregion TestQuerySpcl

        #region FindTheDBwithMeterID
        public string FindTheDBwithMeterID(string meterID, string Database, string dbo_type)
        {
            string tempBatch = string.Empty;
            try
            {
                string Server = "netserver3";

                string query = "SELECT batch from "+ dbo_type + ".Meter where MeterID = '"+ meterID + "' ";

                string ConnectionString = "Server=" + Server + "; Database=" + Database + "; UId=" + "power" + "; Password=" + "power" + ";";

                this.dt = DatabaseQueries.ExecuteQuery(query, ConnectionString);
                if (this.dt.Rows.Count <= 0)
                    return "NoData";

                foreach (DataRow dr in dt.Rows)
                {
                    tempBatch = string.Empty + DatabaseQueries.CheckForNull<dynamic>(dr["Batch"]);
                }
                return tempBatch;
            }
            catch
            {
                return "NoData";
            }
        }
        #endregion FindTheDBwithMeterID

        #region SendSQLRaw
        public DataTable SendSQLRaw(string CommandSQL, string Database)
        {
            string tempBatch = string.Empty;
            try
            {
                string Server = "netserver3";

                string query = CommandSQL;

                string ConnectionString = "Server=" + Server + "; Database=" + Database + "; UId=" + "power" + "; Password=" + "power" + ";";

                this.dt = DatabaseQueries.ExecuteQuery(query, ConnectionString);

                return dt;
            }
            catch
            {
                return dt;
            }
        }
        #endregion SendSQLRaw
    }
    #endregion SQLManager Class
}

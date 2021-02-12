using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

//using VestaDLL;

namespace ExcelReadingApp
{
    [Serializable()]
    public class User
    {
        #region Properties

        public string Server { get; set; }
        //public string Type { get; set; }
        public string Database { get; set; }
        public string DBOwner { get; set; }
        public string ConnectionString { get; set; }
        //public bool UseCustomFields { get; set; }
        public Credentials DBCredentials { get; set; }      // credentials for database
        public Credentials SQLCredentials { get; set; }     // credentials for SQL server or SQL express
       

        #endregion Properties

        #region Constructors

        public User()
        {
            this.Server = string.Empty;
            //this.Type = Noun.SQL;
            this.Database = string.Empty;
            this.DBOwner = string.Empty;
            this.ConnectionString = string.Empty;
            //this.UseCustomFields = false;
            this.DBCredentials = new Credentials();
            this.SQLCredentials = new Credentials();
        }

        #endregion Constructors

        #region Methods

        public void SetConnectionString()
        {
            if (this.Server == string.Empty || this.Database == string.Empty)
                return;

            this.ConnectionString =
                "Server=" + this.Server + "; Database=" + this.Database + "; UId=" + this.SQLCredentials.UserID + "; Password=" + this.SQLCredentials.Password + ";";
        }

        public void SetConnectionString(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        public void SetConnectionString(string server, string database)
        {
            this.ConnectionString =
                "Server=" + server + "; Database=" + database + "; " + "Integrated Security=SSPI;";
            //"Server=" + server + "; Database=" + database + "; Trusted_Connection=" + true;
        }

        public void SetConnectionString(string server, string database, string userid, string password)
        {
            this.ConnectionString =
                "Server=" + server + "; Database=" + database + "; UId=" + userid + "; Password=" + password + ";";
            //"Server=" + server + "; Database=" + database + "; Trusted_Connection=" + true;
        }

        #endregion Methods
    }

    public static class UserManager
    {
        #region Save

        public static void Save(User user, string filePath)
        {
            // save machine state object to XML file using our ObjectXMLSerializer class...
            try
            {
                ObjectXMLSerializer<User>.Save(user, filePath);

                //VestaDLL.Utilities.GrantAccess(filePath);
            }

            catch (Exception ex)
            {
                MessageBox.Show("Unable to save auto login,\r\n"+ex.Message,"Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);//modified from the original code

                //MessageBox.Show("Unable to save auto login to: " + Environment.NewLine + Environment.NewLine + "'" + filePath + "'" + Environment.NewLine + Environment.NewLine +
                //ex.Message + Environment.NewLine + ex.Source + Environment.NewLine + ex.StackTrace, AssemblyInfo.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion Save

        #region Load

        public static User Load(string filePath)
        {
            User user = null;

            if (!File.Exists(filePath))
                return user;

            // Load the user object from the XML file using our custom class...
            user = ObjectXMLSerializer<User>.Load(filePath);

            if (user == null)
            {
                MessageBox.Show("Unable to load auto login from file.");
            }

            return user;
        }

        #endregion Load

        #region Update User Log

        //public static bool UpdateUserLog(string action)
        //{
        //    if (!SQLManager.UserLogTable)
        //        return false;

        //    if (StatusInformation.User.SQLCredentials.UserID == string.Empty)
        //        return false;

        //    if (StatusInformation.User.ConnectionString == string.Empty)
        //        return false;

        //    try
        //    {
        //        StringBuilder sb = new StringBuilder();

        //        sb.Append("INSERT INTO ")
        //          .Append(StatusInformation.User.DBOwner)
        //          .Append(".UserLog")
        //          .Append(" (")
        //          .Append(StatusInformation.User.DBOwner).Append(".").Append("UserLog.Timestamp")
        //          .Append(", ").Append(StatusInformation.User.DBOwner).Append(".").Append("UserLog.UserID")
        //          .Append(", ").Append(StatusInformation.User.DBOwner).Append(".").Append("UserLog.Action")
        //          .Append(", ").Append(StatusInformation.User.DBOwner).Append(".").Append("UserLog.Server")
        //          .Append(", ").Append(StatusInformation.User.DBOwner).Append(".").Append("UserLog.DB")

        //          .Append(") VALUES (")
        //          .Append("@timestamp")
        //          .Append(", @userid")
        //          .Append(", @action")
        //          .Append(", @server")
        //          .Append(", @db")
        //          .Append(")");

        //        VestaDLL.DebugManager.UpdateLog("User (UpdateUserLog)", "Query=" + sb.ToString());

        //        SqlConnection connection = new SqlConnection(StatusInformation.User.ConnectionString);
        //        SqlCommand command = new SqlCommand(sb.ToString(), connection);

        //        SQLManager.LoadParameter(command, "@timestamp", SqlDbType.DateTime, DateTime.Now);
        //        SQLManager.LoadParameter(command, "@userid", SqlDbType.NChar, StatusInformation.User.DBCredentials.UserID);
        //        SQLManager.LoadParameter(command, "@action", SqlDbType.NChar, action);
        //        SQLManager.LoadParameter(command, "@server", SqlDbType.NChar, StatusInformation.User.Server);
        //        SQLManager.LoadParameter(command, "@db", SqlDbType.NChar, StatusInformation.User.Database);

        //        connection.Open();
        //        command.ExecuteNonQuery();

        //        //+VestaDLL.DebugManager.UpdateLog("Meter (InsertMeterTestHistoryRecord)", "Record inserted successfully");

        //        return false;
        //    }

        //    catch (Exception e)
        //    {
        //        Utilities.ShowMessageBox(
        //            e.Message + Environment.NewLine + e.Source + Environment.NewLine + Environment.NewLine + e.StackTrace,
        //            "Program Exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

        //        //+VestaDLL.DebugManager.UpdateLog("Meter (InsertMeterTestHistoryRecords)", "Record insertion failed");

        //        return true;
        //    }
        //}

        #endregion Update User Log
    }
}

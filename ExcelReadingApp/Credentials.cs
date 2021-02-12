using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadingApp
{
    [Serializable()]
    public class Credentials
    {
        #region Properties

        public string UserID { get; set; }
        public string Password { get; set; }

        #endregion Properties

        #region Constructors

        public Credentials()
        {
            this.UserID = string.Empty;
            this.Password = string.Empty;
        }

        #endregion Constructors
    }
}
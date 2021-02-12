using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReadingApp
{
    public partial class Authentication : Form
    {
        public Authentication()
        {
            InitializeComponent();
        }

        public bool ChecktheUserPass()
        {
            Form1 F1 = new Form1();

            if (string.Equals(F1.CheckTheUsername(textBox_P_user.Text.ToUpper()), "OK"))
            {
                if (string.Equals(textBox_P_pin.Text, "1234"))
                {
                    return true;
                }
            }
            else { return false; }
            return false;
        }
    }
}

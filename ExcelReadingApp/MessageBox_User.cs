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
    public partial class MessageBox_User : Form
    {
        public MessageBox_User()
        {
            InitializeComponent();
        }


        public void MB_TextDisplay(string textToDisplay)
        {
            richTextBox_MBU.Text = textToDisplay;
        }
        public void MB_TextAppend(string textToDisplay)
        {
            richTextBox_MBU.AppendText(textToDisplay);// = textToDisplay;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
        }
    }
}

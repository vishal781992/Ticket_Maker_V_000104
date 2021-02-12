using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace ExcelReadingApp
{
    public partial class DirectoryRefresh : Form
    {
        private BackgroundWorker BGW;
        public DirectoryRefresh()
        {
            InitializeComponent();
            BGW = new BackgroundWorker();
            BGW.WorkerReportsProgress = true;
            BGW.WorkerSupportsCancellation = true;
            BGW.DoWork += BGW1_DoWork;
            BGW.RunWorkerCompleted += BGW1_RunWorkerCompleted;
            BGW.ProgressChanged += BGW1_ProgressChanged;
        }
        //public bool Rambo = false;
        private void DirectoryRefresh_Load(object sender, EventArgs e)
        {
            //progressBar_DirectoryRefresh.Value = 20;
            button1.Visible = false;
            BGW.RunWorkerAsync();
            progressBar_DirectoryRefresh.Maximum = 100;
            //Form1 F1 = new Form1();
            //F1.refresh();

        }

        int count = 0;

        private void BGW1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            //Thread.Sleep(500);
            //worker.ReportProgress(1);
            while (count < 100)
            {
                Thread.Sleep(200);
                worker.ReportProgress(count);
                count = count + 10;
            }
        }

        private void BGW1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            button1.Visible = true;
            this.progressBar_DirectoryRefresh.Text = "100%";
            this.progressBar_DirectoryRefresh.Value = this.progressBar_DirectoryRefresh.Maximum;
        }

        private void BGW1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar_DirectoryRefresh.Value = e.ProgressPercentage;
            this.progressBar_DirectoryRefresh.Text = this.progressBar_DirectoryRefresh.Value + "%";
            this.progressBar_DirectoryRefresh.Refresh();
        }


        public void PB_increment(int value)
        {
            progressBar_DirectoryRefresh.Value = value;
        }

        private void button_Enter_Click(object sender, EventArgs e)
        {
            BGW.RunWorkerAsync();
            progressBar_DirectoryRefresh.Maximum = 100;
            //Form1 F1 = new Form1();
            //F1.refresh();
        }
    }
}

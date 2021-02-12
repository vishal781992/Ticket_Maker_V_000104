namespace ExcelReadingApp
{
    partial class DirectoryRefresh
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.progressBar_DirectoryRefresh = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button_Enter = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // progressBar_DirectoryRefresh
            // 
            this.progressBar_DirectoryRefresh.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuPopup;
            this.progressBar_DirectoryRefresh.Location = new System.Drawing.Point(12, 36);
            this.progressBar_DirectoryRefresh.Name = "progressBar_DirectoryRefresh";
            this.progressBar_DirectoryRefresh.Size = new System.Drawing.Size(382, 31);
            this.progressBar_DirectoryRefresh.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(91, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(219, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Wait, The Directory is Being Refreshed!";
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(235, 101);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button_Enter
            // 
            this.button_Enter.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.button_Enter.Location = new System.Drawing.Point(75, 101);
            this.button_Enter.Name = "button_Enter";
            this.button_Enter.Size = new System.Drawing.Size(75, 23);
            this.button_Enter.TabIndex = 3;
            this.button_Enter.Text = "Enter";
            this.button_Enter.UseVisualStyleBackColor = true;
            this.button_Enter.Click += new System.EventHandler(this.button_Enter_Click);
            // 
            // DirectoryRefresh
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 136);
            this.ControlBox = false;
            this.Controls.Add(this.button_Enter);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar_DirectoryRefresh);
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.Name = "DirectoryRefresh";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Directory Refresh";
            this.TopMost = true;
            this.UseWaitCursor = true;
            this.Load += new System.EventHandler(this.DirectoryRefresh_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar_DirectoryRefresh;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button_Enter;
    }
}
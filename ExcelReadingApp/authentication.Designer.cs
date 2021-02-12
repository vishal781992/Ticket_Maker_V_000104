namespace ExcelReadingApp
{
    partial class Authentication
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_P_user = new System.Windows.Forms.TextBox();
            this.textBox_P_pin = new System.Windows.Forms.TextBox();
            this.button_P_submit = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(18, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "User:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(18, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Pin Code";
            // 
            // textBox_P_user
            // 
            this.textBox_P_user.Location = new System.Drawing.Point(78, 17);
            this.textBox_P_user.Name = "textBox_P_user";
            this.textBox_P_user.Size = new System.Drawing.Size(218, 20);
            this.textBox_P_user.TabIndex = 2;
            // 
            // textBox_P_pin
            // 
            this.textBox_P_pin.Location = new System.Drawing.Point(78, 52);
            this.textBox_P_pin.Name = "textBox_P_pin";
            this.textBox_P_pin.PasswordChar = '#';
            this.textBox_P_pin.Size = new System.Drawing.Size(218, 20);
            this.textBox_P_pin.TabIndex = 3;
            // 
            // button_P_submit
            // 
            this.button_P_submit.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button_P_submit.Location = new System.Drawing.Point(221, 78);
            this.button_P_submit.Name = "button_P_submit";
            this.button_P_submit.Size = new System.Drawing.Size(75, 23);
            this.button_P_submit.TabIndex = 4;
            this.button_P_submit.Text = "Submit";
            this.button_P_submit.UseVisualStyleBackColor = true;
            // 
            // button_cancel
            // 
            this.button_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_cancel.Location = new System.Drawing.Point(21, 78);
            this.button_cancel.Name = "button_cancel";
            this.button_cancel.Size = new System.Drawing.Size(75, 23);
            this.button_cancel.TabIndex = 5;
            this.button_cancel.Text = "Cancel";
            this.button_cancel.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(75, 104);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(190, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "If cant access, email admin for access.";
            // 
            // Authentication
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.ClientSize = new System.Drawing.Size(322, 128);
            this.ControlBox = false;
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.button_P_submit);
            this.Controls.Add(this.textBox_P_pin);
            this.Controls.Add(this.textBox_P_user);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Name = "Authentication";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Authentication";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox_P_user;
        private System.Windows.Forms.TextBox textBox_P_pin;
        private System.Windows.Forms.Button button_P_submit;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.Label label3;
    }
}
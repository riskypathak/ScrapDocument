namespace ScrapDocument
{
    partial class ScrapDocument
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
            this.grpExtract = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbTemplate = new System.Windows.Forms.ComboBox();
            this.btnStart = new System.Windows.Forms.Button();
            this.SelectDirectory = new System.Windows.Forms.Button();
            this.txtDirectory = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.grpLog = new System.Windows.Forms.GroupBox();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.fbdDirectory = new System.Windows.Forms.FolderBrowserDialog();
            this.fbdSaveFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.grpExtract.SuspendLayout();
            this.grpLog.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpExtract
            // 
            this.grpExtract.Controls.Add(this.label2);
            this.grpExtract.Controls.Add(this.cmbTemplate);
            this.grpExtract.Controls.Add(this.btnStart);
            this.grpExtract.Controls.Add(this.SelectDirectory);
            this.grpExtract.Controls.Add(this.txtDirectory);
            this.grpExtract.Controls.Add(this.label1);
            this.grpExtract.Location = new System.Drawing.Point(27, 12);
            this.grpExtract.Name = "grpExtract";
            this.grpExtract.Size = new System.Drawing.Size(364, 128);
            this.grpExtract.TabIndex = 0;
            this.grpExtract.TabStop = false;
            this.grpExtract.Text = "Extract";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Template";
            // 
            // cmbTemplate
            // 
            this.cmbTemplate.Location = new System.Drawing.Point(61, 70);
            this.cmbTemplate.Name = "cmbTemplate";
            this.cmbTemplate.Size = new System.Drawing.Size(182, 21);
            this.cmbTemplate.TabIndex = 4;
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(260, 68);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 3;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // SelectDirectory
            // 
            this.SelectDirectory.Location = new System.Drawing.Point(260, 23);
            this.SelectDirectory.Name = "SelectDirectory";
            this.SelectDirectory.Size = new System.Drawing.Size(75, 23);
            this.SelectDirectory.TabIndex = 2;
            this.SelectDirectory.Text = "Select";
            this.SelectDirectory.UseVisualStyleBackColor = true;
            this.SelectDirectory.Click += new System.EventHandler(this.Select_Click);
            // 
            // txtDirectory
            // 
            this.txtDirectory.Location = new System.Drawing.Point(61, 23);
            this.txtDirectory.Name = "txtDirectory";
            this.txtDirectory.ReadOnly = true;
            this.txtDirectory.Size = new System.Drawing.Size(182, 20);
            this.txtDirectory.TabIndex = 1;
            this.txtDirectory.TextChanged += new System.EventHandler(this.txtDirectory_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Directory";
            // 
            // grpLog
            // 
            this.grpLog.Controls.Add(this.txtLog);
            this.grpLog.Location = new System.Drawing.Point(27, 166);
            this.grpLog.Name = "grpLog";
            this.grpLog.Size = new System.Drawing.Size(364, 181);
            this.grpLog.TabIndex = 2;
            this.grpLog.TabStop = false;
            this.grpLog.Text = "Log";
            // 
            // txtLog
            // 
            this.txtLog.Location = new System.Drawing.Point(6, 26);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtLog.Size = new System.Drawing.Size(329, 140);
            this.txtLog.TabIndex = 6;
            this.txtLog.TextChanged += new System.EventHandler(this.txtLog_TextChanged);
            // 
            // fbdDirectory
            // 
            this.fbdDirectory.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.fbdDirectory.ShowNewFolderButton = false;
            // 
            // ScrapDocument
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 362);
            this.Controls.Add(this.grpLog);
            this.Controls.Add(this.grpExtract);
            this.Name = "ScrapDocument";
            this.Text = "ScrapDocument";
            this.grpExtract.ResumeLayout(false);
            this.grpExtract.PerformLayout();
            this.grpLog.ResumeLayout(false);
            this.grpLog.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpExtract;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button SelectDirectory;
        private System.Windows.Forms.TextBox txtDirectory;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpLog;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.FolderBrowserDialog fbdDirectory;
        private System.Windows.Forms.FolderBrowserDialog fbdSaveFolder;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbTemplate;
    }
}


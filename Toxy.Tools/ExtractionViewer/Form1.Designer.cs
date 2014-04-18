namespace ExtractionViewer
{
    partial class Form1
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.excelDataSetViewerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.btnReopen = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.tbExtension = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbParserType = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.cbSheets = new System.Windows.Forms.ComboBox();
            this.btnSelectSheet = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1086, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(39, 21);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(108, 22);
            this.openToolStripMenuItem.Text = "&Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(108, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.excelDataSetViewerToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(52, 21);
            this.toolsToolStripMenuItem.Text = "&Tools";
            // 
            // excelDataSetViewerToolStripMenuItem
            // 
            this.excelDataSetViewerToolStripMenuItem.Name = "excelDataSetViewerToolStripMenuItem";
            this.excelDataSetViewerToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.excelDataSetViewerToolStripMenuItem.Text = "DataGrid Viewer for Excel/CSV";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem1});
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(47, 21);
            this.aboutToolStripMenuItem.Text = "&Help";
            // 
            // aboutToolStripMenuItem1
            // 
            this.aboutToolStripMenuItem1.Name = "aboutToolStripMenuItem1";
            this.aboutToolStripMenuItem1.Size = new System.Drawing.Size(111, 22);
            this.aboutToolStripMenuItem1.Text = "&About";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 25);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.comboBox1);
            this.splitContainer1.Panel2.Controls.Add(this.btnReopen);
            this.splitContainer1.Panel2.Controls.Add(this.groupBox1);
            this.splitContainer1.Size = new System.Drawing.Size(1086, 465);
            this.splitContainer1.SplitterDistance = 785;
            this.splitContainer1.TabIndex = 5;
            // 
            // comboBox1
            // 
            this.comboBox1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.comboBox1.Items.AddRange(new object[] {
            "UTF-8",
            "UTF-7",
            "GB2312",
            "GBK",
            "BIG5",
            "UTF-32",
            "Shift-JIS"});
            this.comboBox1.Location = new System.Drawing.Point(12, 14);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(179, 20);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.Text = "UTF-8";
            // 
            // btnReopen
            // 
            this.btnReopen.Location = new System.Drawing.Point(210, 11);
            this.btnReopen.Name = "btnReopen";
            this.btnReopen.Size = new System.Drawing.Size(75, 23);
            this.btnReopen.TabIndex = 6;
            this.btnReopen.Text = "Reopen";
            this.btnReopen.UseVisualStyleBackColor = true;
            this.btnReopen.Click += new System.EventHandler(this.btnReopen_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.tbPath);
            this.groupBox1.Controls.Add(this.tbExtension);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.tbParserType);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(3, 40);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(282, 414);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "File Extraction Info";
            // 
            // tbPath
            // 
            this.tbPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbPath.Location = new System.Drawing.Point(6, 114);
            this.tbPath.Multiline = true;
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(267, 85);
            this.tbPath.TabIndex = 5;
            // 
            // tbExtension
            // 
            this.tbExtension.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbExtension.Location = new System.Drawing.Point(96, 65);
            this.tbExtension.Name = "tbExtension";
            this.tbExtension.Size = new System.Drawing.Size(129, 21);
            this.tbExtension.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(7, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 12);
            this.label3.TabIndex = 3;
            this.label3.Text = "Extension: ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 99);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "Path:";
            // 
            // tbParserType
            // 
            this.tbParserType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tbParserType.Location = new System.Drawing.Point(96, 29);
            this.tbParserType.Name = "tbParserType";
            this.tbParserType.Size = new System.Drawing.Size(129, 21);
            this.tbParserType.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Parser Type: ";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSelectSheet);
            this.panel1.Controls.Add(this.cbSheets);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Location = new System.Drawing.Point(9, 222);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(264, 64);
            this.panel1.TabIndex = 7;
            this.panel1.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 10);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "Sheet: ";
            // 
            // cbSheets
            // 
            this.cbSheets.FormattingEnabled = true;
            this.cbSheets.Location = new System.Drawing.Point(58, 7);
            this.cbSheets.Name = "cbSheets";
            this.cbSheets.Size = new System.Drawing.Size(203, 20);
            this.cbSheets.TabIndex = 8;
            // 
            // btnSelectSheet
            // 
            this.btnSelectSheet.Location = new System.Drawing.Point(198, 33);
            this.btnSelectSheet.Name = "btnSelectSheet";
            this.btnSelectSheet.Size = new System.Drawing.Size(64, 23);
            this.btnSelectSheet.TabIndex = 9;
            this.btnSelectSheet.Text = "Select";
            this.btnSelectSheet.UseVisualStyleBackColor = true;
            this.btnSelectSheet.Click += new System.EventHandler(this.btnSelectSheet_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1086, 490);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Toxy Extraction Viewer";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem excelDataSetViewerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button btnReopen;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tbParserType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.TextBox tbExtension;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnSelectSheet;
        private System.Windows.Forms.ComboBox cbSheets;
        private System.Windows.Forms.Label label4;
    }
}


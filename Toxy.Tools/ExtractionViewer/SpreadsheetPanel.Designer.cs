using System.Windows.Forms;
using unvell.ReoGrid;
namespace ExtractionViewer
{
    partial class SpreadsheetPanel
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // SpreadsheetPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Name = "SpreadsheetPanel";
            this.Size = new System.Drawing.Size(640, 322);
            this.ResumeLayout(false);

            this.reoGridControl1 = new ReoGridControl();
            this.reoGridControl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.reoGridControl1.CellContextMenuStrip = null;
            this.reoGridControl1.ColCount = 100;
            this.reoGridControl1.ColHeadContextMenuStrip = null;
            this.reoGridControl1.Location = new System.Drawing.Point(189, 243);
            this.reoGridControl1.Name = "reoGridControl1";
            this.reoGridControl1.RowCount = 200;
            this.reoGridControl1.RowHeadContextMenuStrip = null;
            this.reoGridControl1.Script = null;
            this.reoGridControl1.Dock = DockStyle.Fill;
            this.reoGridControl1.TabIndex = 0;
            this.reoGridControl1.Text = "reoGridControl1";
            this.Controls.Add(this.reoGridControl1);
        }

        #endregion

        private unvell.ReoGrid.ReoGridControl reoGridControl1;
    }
}

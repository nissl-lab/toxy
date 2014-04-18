using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Toxy;

namespace ExtractionViewer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        string filepath = null;

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter = "All Supported Files |*.csv;*.txt;*.xls;*.xlsx;*.docx;*.rtf;*.eml;*.xml;*.html;*.htm;*.pdf;*.vcard";
            dialog.Filter += "|Comma Seperated Files (*.csv)|*.csv";
            dialog.Filter += "|Text Files (*.txt)|*.txt";
            dialog.Filter += "|All Excel Files|*.xls;*.xlsx";
            dialog.Filter += "|Rich Text Files (*.rtf)|*.rtf";
            dialog.Filter += "|Word 2007-2013 Files (*.docx)|*.docx";
            dialog.Filter += "|Business Card Files (*.vcard)|*.vcard";
            dialog.Filter += "|Email Files (*.eml)|*.eml";
            dialog.Filter += "|Html Files (*.html, *.htm)|*.html;*.htm";
            dialog.Filter += "|XML Files (*.xml)|*.xml";
            dialog.Filter += "|Adobe PDF Files (*.pdf)|*.pdf";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filepath = dialog.FileName;
            }

            OpenFile(filepath, comboBox1.Text);
        }


        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.DataGridView dataGridView1;

        private void AppendRichTextBox()
        {
            
            
            if (richTextBox1 == null)
            {
                this.richTextBox1 = new System.Windows.Forms.RichTextBox();
                this.richTextBox1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.richTextBox1.Location = new System.Drawing.Point(0, 0);
                this.richTextBox1.Name = "richTextBox1";
                this.richTextBox1.TabIndex = 1;
                this.richTextBox1.Text = "";
                this.splitContainer1.Panel1.Controls.Add(this.richTextBox1);
            }
            this.richTextBox1.Clear();
            this.richTextBox1.Visible = true;
            if(this.dataGridView1!=null)
                this.dataGridView1.Visible = false;
        }
        private void AppendDataGridView()
        {
            if (dataGridView1 == null)
            {
                this.dataGridView1 = new System.Windows.Forms.DataGridView();
                ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
                this.dataGridView1.AllowUserToAddRows = false;
                this.dataGridView1.AllowUserToDeleteRows = false;
                this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
                var headerstyle = new DataGridViewCellStyle();
                headerstyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                this.dataGridView1.ColumnHeadersDefaultCellStyle = headerstyle;
                this.dataGridView1.Name = "dataGridView1";
                this.dataGridView1.ReadOnly = true;
                this.dataGridView1.RowTemplate.Height = 23;
                this.dataGridView1.TabIndex = 0;

                this.splitContainer1.Panel1.Controls.Add(this.dataGridView1);
            }

            if (richTextBox1 != null)
                this.richTextBox1.Visible = false;
            this.dataGridView1.Visible = true;
        }
        ToxySpreadsheet ss = null;
        
        private void OpenFile(string filepath, string encoding)
        {

            if (string.IsNullOrWhiteSpace(filepath))
            {
                tbPath.Clear();
                return;
            }
            
            tbPath.Text = filepath;
            FileInfo fi = new FileInfo(filepath);
            ParserContext context = new ParserContext(filepath);
            context.Encoding = Encoding.GetEncoding(encoding);
            string extension = fi.Extension;
            tbExtension.Text = extension;



            panel1.Visible = false;


            switch (extension)
            {
                case ".txt":
                case ".html":
                case ".htm":
                    AppendRichTextBox();
                    var tparser = ParserFactory.CreateText(context);
                    richTextBox1.Text = tparser.Parse();
                    tbParserType.Text = tparser.GetType().Name;
                    break;
                case ".rtf":
                case ".docx":
                    AppendRichTextBox();
                    IDocumentParser docparser = ParserFactory.CreateDocument(context);
                    ToxyDocument doc = docparser.Parse();
                    tbParserType.Text = docparser.GetType().Name;
                    richTextBox1.Text = doc.ToString();
                    break;
                case ".csv":
                case ".xlsx":
                case ".xls":
                    AppendDataGridView();
                    ISpreadsheetParser ssparser = ParserFactory.CreateSpreadsheet(context);
                    ss = ssparser.Parse();
                    DataSet ds = ss.ToDataSet();
                    dataGridView1.DataSource = ds.Tables[0].DefaultView;

                    cbSheets.Items.Clear();
                    foreach (var table in ss.Tables)
                    {
                        cbSheets.Items.Add(table.Name);
                    }
                    cbSheets.SelectedIndex = 0;
                    panel1.Visible = true;
                    break;
                case ".vcard":
                    AppendRichTextBox();
                    var vparser = ParserFactory.CreateVCard(context);
                    ToxyBusinessCards vcards = vparser.Parse();
                    tbParserType.Text = vparser.GetType().Name;
                    richTextBox1.Text = vcards.ToString();
                    break;
                default:
                    AppendRichTextBox();
                    richTextBox1.Text = "Unknown document";
                    tbParserType.Text = "";
                    break;
            }
            
        }

        private void btnReopen_Click(object sender, EventArgs e)
        {
            OpenFile(filepath, comboBox1.Text);
        }

        private void btnSelectSheet_Click(object sender, EventArgs e)
        {
            if (ss == null)
            {
                return;
            }
            var table= ss[cbSheets.Text];
            if(table==null)
                return;
                dataGridView1.DataSource = table.ToDataTable().DefaultView;
        }

    }
}

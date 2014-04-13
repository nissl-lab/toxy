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

            switch (extension)
            {
                case ".txt":
                case ".html":
                case ".htm":
                    var tparser = ParserFactory.CreateText(context);
                    richTextBox1.Text = tparser.Parse();
                    tbParserType.Text = tparser.GetType().Name;
                    break;
                case ".rtf":
                case ".docx":
                    IDocumentParser docparser = ParserFactory.CreateDocument(context);
                    ToxyDocument doc = docparser.Parse();
                    tbParserType.Text = docparser.GetType().Name;
                    richTextBox1.Text = doc.ToString();
                    break;
                case ".csv":
                case ".xlsx":
                case ".xls":
                    ISpreadsheetParser ssparser = ParserFactory.CreateSpreadsheet(context);
                    ToxySpreadsheet ss = ssparser.Parse();
                    tbParserType.Text = ssparser.GetType().Name;
                    richTextBox1.Text = ss.ToString();
                    break;
                case ".vcard":
                    var vparser = ParserFactory.CreateVCard(context);
                    ToxyBusinessCards vcards = vparser.Parse();
                    tbParserType.Text = vparser.GetType().Name;
                    richTextBox1.Text = vcards.ToString();
                    break;
                default:
                    richTextBox1.Text = "Unknown document";
                    tbParserType.Text = "";
                    break;
            }
            
        }

        private void btnReopen_Click(object sender, EventArgs e)
        {
            OpenFile(filepath, comboBox1.Text);
        }

    }
}

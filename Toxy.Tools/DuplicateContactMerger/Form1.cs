using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Toxy;
using Toxy.Parsers;

namespace DuplicateContactMerger
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        void FillListView(ToxyBusinessCards bcs, bool includeEmptyContacts, bool includeDuplicateName)
        {
            listView1.Items.Clear();
            int total = 0;
            bcs.IncludeEmptyContact = includeEmptyContacts;

            Dictionary<string, ToxyBusinessCard> tbcs = new Dictionary<string, ToxyBusinessCard>();

            foreach (var card in bcs.Cards)
            {
                ListViewItem item = new ListViewItem();
                item.Text = card.Name.FullName;
                string fax = null;
                int phoneCount = 0;
                bool isEmpty = true;

                if (!includeDuplicateName)
                {
                    if (!tbcs.ContainsKey(card.Name.FullName))
                    {
                        tbcs.Add(card.Name.FullName, card);
                    }
                    else
                    {
                        continue;
                    }
                }

                foreach (var contact in card.Contacts)
                {
                    if (contact.Name == "Cellular" || contact.Name == "Home" || contact.Name == "Work" || contact.Name == "Voice")
                    {
                        phoneCount++;
                        item.SubItems.Add(contact.Value);
                        isEmpty = false;
                    }
                    else if (contact.Name == "Internet" || contact.Name == "Url-Default")
                    {
                        //item.SubItems.Add(contact.Value);
                    }
                    else if (contact.Name.StartsWith("Fax,"))
                    {
                        fax = contact.Value;
                        isEmpty = false;
                    }
                    else
                    {
                        string x = contact.Name;
                    }
                }
                if (!string.IsNullOrEmpty(fax))
                {
                    while (3 - phoneCount > 0)
                    {
                        item.SubItems.Add("");
                        phoneCount++;
                    }
                    item.SubItems.Add(fax);
                }
                if (includeEmptyContacts)
                {
                    listView1.Items.Add(item);
                    total++;
                }
                else if (!isEmpty)
                {
                    total++;
                    listView1.Items.Add(item);
                }

            }
            label2.Text = total.ToString();
            
        }
        ToxyBusinessCards bcs=null;
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Filter = "Business Card Files (*.vcf)|*.vcf";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string path= dialog.FileName;
                ParserContext context=new ParserContext(path);
                VCardDocumentParser vparser=ParserFactory.CreateVCard(context);
                bcs = vparser.Parse();

                FillListView(bcs, checkBox1.Checked, checkBox2.Checked);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb1 = (CheckBox)sender;
            if (bcs != null)
            {
                FillListView(bcs, checkBox1.Checked, checkBox2.Checked);
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox cb1 = (CheckBox)sender;
            if (bcs != null)
            {
                FillListView(bcs, checkBox1.Checked, checkBox2.Checked);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Business Card Files (*.vcf)|*.vcf";
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string savePath = sfd.FileName;
            }
        }
    }
}

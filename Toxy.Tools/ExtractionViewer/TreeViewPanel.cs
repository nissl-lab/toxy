using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExtractionViewer
{
    public partial class TreeViewPanel : UserControl
    {
        public TreeViewPanel()
        {
            InitializeComponent();
        }

        public TreeView Tree
        {
            get { return treeView1; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.treeView1.BeginUpdate();
            this.treeView1.ExpandAll();
            this.treeView1.EndUpdate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.treeView1.BeginUpdate();
            this.treeView1.CollapseAll();
            this.treeView1.EndUpdate();
        }
    }   
}

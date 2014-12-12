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
    public partial class GridPanel : UserControl
    {
        public GridPanel()
        {
            InitializeComponent();
        }

        public DataGridView GridView
        {
            get { return this.dataGridView1; }
        }
    }
}

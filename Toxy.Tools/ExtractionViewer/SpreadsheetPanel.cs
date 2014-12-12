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
    public partial class SpreadsheetPanel : UserControl
    {
        public SpreadsheetPanel()
        {
            InitializeComponent();
        }

        public unvell.ReoGrid.ReoGridControl ReoGridControl
        {
            get { return this.reoGridControl1; }
        }
    }
}

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
    public partial class RichTextBoxPanel : UserControl
    {
        public RichTextBoxPanel()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            richTextBox1.Clear();
        }

        public override string Text
        {
            get { return this.richTextBox1.Text; }
            set { this.richTextBox1.Text = value; }
        }
    }
}

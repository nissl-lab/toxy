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
    public partial class PropertyListPanel : UserControl
    {
        public PropertyListPanel()
        {
            InitializeComponent();
        }

        public void AddItem(string itemName, string value)
        {
            var item = listView1.Items.Add(itemName);
            item.SubItems.Add(value);
        }

        public void Clear()
        {
            this.listView1.Items.Clear();
        }
    }
}

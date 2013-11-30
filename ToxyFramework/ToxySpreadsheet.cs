using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyRow
    {
        public ToxyRow()
        {
            Cells = new List<string>();
        }
        public List<string> Cells {get;set;}
    }

    public class ToxySpreadsheet
    {
        public ToxySpreadsheet()
        {
            this.Headers = new List<string>();
            this.Rows = new List<ToxyRow>();
            this.Properties = new List<IProperty>();
        }
        public List<string> Headers { get; set; }
        public List<ToxyRow> Rows { get; set; }
        public int TotalSheetNumber { get; set; }
        public List<IProperty> Properties { get; set; }
    }
}

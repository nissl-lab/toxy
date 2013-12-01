using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Toxy
{
    public class ToxyTable
    {
        public ToxyTable()
        {
            this.ColumnHeaders = new List<string>();
            this.Rows = new List<ToxyRow>();
        }
        public string Name { get; set; }
        public string PageHeader { get; set; }
        public string PageFooter { get; set; }

        public List<string> ColumnHeaders { get; set; }
        public List<ToxyRow> Rows { get; set; }        
    }

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
            this.Tables = new List<ToxyTable>();
            this.Properties = new Dictionary<string, object>();
        }
        public List<ToxyTable> Tables { get; set; }
        public Dictionary<string, object> Properties { get; set; }

        public DataSet ToDataSet()
        {
            throw new NotImplementedException();
        }
    }
}

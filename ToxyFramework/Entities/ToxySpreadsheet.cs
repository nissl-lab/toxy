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
            this.ColumnHeaders = new ToxyRow(0);
            this.Rows = new List<ToxyRow>();
            this.LastColumnIndex = -1;
        }
        public bool HasHeader 
        {
            get { return this.ColumnHeaders.Cells.Count > 0; }
        }
        public string Name { get; set; }
        public string PageHeader { get; set; }
        public string PageFooter { get; set; }
        public int LastRowIndex { get; set; }
        public int LastColumnIndex { get; set; }

        public ToxyRow ColumnHeaders { get; set; }

        public List<ToxyRow> Rows { get; set; }

        public DataTable ToDataTable()
        {

            int lastCol = 0;
            DataTable dt = new DataTable(this.Name);
            if (HasHeader)
            {
                foreach (var header in ColumnHeaders.Cells)
                {
                    var col = new DataColumn(header.Value);
                    dt.Columns.Add(col);
                    lastCol++;
                }
                for (int j = dt.Columns.Count; this.LastColumnIndex > 0 && j <= this.LastColumnIndex; j++)
                {
                    dt.Columns.Add(string.Empty);
                    lastCol++;
                }
            }
            int lastRow=0;
            foreach (var row in this.Rows)
            {
                DataRow drow = null;
                if (lastRow < row.RowIndex)
                {
                    drow = dt.NewRow();
                    while (lastRow < row.RowIndex)
                    {
                        dt.Rows.Add(drow);
                        drow = dt.NewRow();
                        lastRow++;
                    }
                }
                else
                {
                    drow = dt.NewRow();
                }

                while (lastCol < row.LastCellIndex)
                {
                    dt.Columns.Add("Column" + lastCol);
                    lastCol++;
                }

                foreach(var cell in row.Cells)
                {                    
                    drow[cell.CellIndex] = cell.Value;   //no comment included
                }
                if (drow == null)
                {
                    drow = dt.NewRow();
                }
                dt.Rows.Add(drow);
                lastRow++;
            }
            return dt;
        }
    }

    public class ToxyRow
    {
        public int RowIndex { get; set; }
        public int LastCellIndex { get; set; }
        public ToxyRow(int rowIndex)
        {
            this.RowIndex = rowIndex;
            this.Cells = new List<ToxyCell>();
        }
        public List<ToxyCell> Cells
        {
            get;
            set;
        }
    }
    public class ToxyCell
    {
        public ToxyCell(int cellIndex, string value)
        {
            if (value == null)
            {
                this.Value = string.Empty;
            }
            else
            {
                this.Value = value;
            }
            this.CellIndex = cellIndex;            
        }

        public string Value { get; set; }
        public int CellIndex { get; set; }
        public string Comment { get; set; }
        public override string ToString()
        {
            return this.Value;
        }
    }

    public class ToxySpreadsheet : IToxyProperties
    {
        public ToxySpreadsheet()
        {
            this.Tables = new List<ToxyTable>();
            this.Properties = new Dictionary<string, object>();
        }
        public string Name { get; set; }
        public List<ToxyTable> Tables { get; set; }
        public Dictionary<string, object> Properties { get; set; }


        public DataSet ToDataSet()
        {
            DataSet ds = new DataSet();
            ds.DataSetName = this.Name;
            foreach (var table in this.Tables)
            {
                ds.Tables.Add(table.ToDataTable());
            }
            return ds;
        }
    }
}

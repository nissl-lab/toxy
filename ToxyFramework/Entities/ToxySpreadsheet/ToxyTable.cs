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
            this.Name = string.Empty;
            this.ColumnHeaders = new ToxyRow(0);
            this.Rows = new List<ToxyRow>();
            this.LastColumnIndex = -1;
            this.MergeCells = new List<MergeCellRange>();
        }
        public bool HasHeader
        {
            get { return this.ColumnHeaders.Cells.Count > 0; }
        }

        public List<MergeCellRange> MergeCells { get; set; }
        public string Name { get; set; }
        public string PageHeader { get; set; }
        public string PageFooter { get; set; }
        public int LastRowIndex { get; set; }
        public int LastColumnIndex { get; set; }

        public ToxyRow ColumnHeaders { get; set; }

        public List<ToxyRow> Rows { get; set; }

        private string ExcelColumnFromNumber(int column)
        {
            column++;
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString; 
        }

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
            int lastRow = 0;
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

                while (lastCol <= row.LastCellIndex)
                {
                    dt.Columns.Add(ExcelColumnFromNumber(lastCol));
                    lastCol++;
                }

                foreach (var cell in row.Cells)
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

        public override string ToString()
        {
            return string.Format("[{0}]",this.Name);
        }
    }
}

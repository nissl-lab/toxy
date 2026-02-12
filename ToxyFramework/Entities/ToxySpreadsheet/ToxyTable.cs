using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Linq;

namespace Toxy
{

    public class ToxyTable:ICloneable, IPrettyTable
    {
        public ToxyTable()
        {
            this.Name = string.Empty;
            this.HeaderRows = new List<ToxyRow>();
            this.Rows = new List<ToxyRow>();
            this.LastColumnIndex = -1;
            this.MergeCells = new List<MergeCellRange>();
        }
        public static int HeaderRowCount=1;
        /// <summary>
        /// The sheet has column header (row)
        /// </summary>
        public bool HasHeader
        {
            get { return this.HeaderRows.Count> 0; }
        }
        public int SheetIndex { get; set; }
        public List<MergeCellRange> MergeCells { get; set; }
        public string Name { get; set; }
        public string PageHeader { get; set; }
        public string PageFooter { get; set; }
        public int LastRowIndex { get; set; }
        public int LastColumnIndex { get; set; }

        public List<ToxyRow> HeaderRows { get; set; }

        public List<ToxyRow> Rows { get; set; }

        public ToxyRow this[int i]
        {
            get
            {
                if (HasHeader)
                    return this.Rows.SingleOrDefault(r => r.RowIndex- HeaderRowCount == i);
                else
                    return this.Rows.SingleOrDefault(r => r.RowIndex == i);
            }
        }
        public int Length
        {
            get { return this.Rows.Count; }
        }
        public ToxyRow[] Slice(int start, int length)
        {
            if (start < 0 || this.Rows.Count==0 || this.Rows.Max(t => t.RowIndex) < start + length - 1)
                throw new ArgumentOutOfRangeException();

            var slice =  this.Rows.Where(r=>r.RowIndex>=start && r.RowIndex<=start+length).ToArray();
            return slice;
        }
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

            int rowIndex = 0;
            if (HasHeader)
            {
                foreach (var header in HeaderRows[0].Cells)
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
                rowIndex++;
            }
            foreach (var row in this.Rows)
            {
                DataRow drow = null;
                if (rowIndex < row.RowIndex)
                {
                    drow = dt.NewRow();
                    while (rowIndex < row.RowIndex)
                    {
                        dt.Rows.Add(drow);
                        drow = dt.NewRow();
                        rowIndex++;
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
                dt.Rows.Add(drow);
                rowIndex++;
            }
            return dt;
        }

        public override string ToString()
        {
            return string.Format("[{0}]",this.Name);
        }

        public object Clone()
        {
            ToxyTable newtt = new ToxyTable();
            newtt.PageFooter = this.PageFooter;
            newtt.PageHeader = this.PageHeader;
            newtt.Name = this.Name;
            newtt.LastColumnIndex = this.LastColumnIndex;
            newtt.LastRowIndex = this.LastRowIndex;
            foreach(ToxyRow row in this.Rows)
            {
                newtt.Rows.Add(row.Clone() as ToxyRow);
            }
            foreach (ToxyRow header in this.HeaderRows)
            {
                newtt.HeaderRows.Add(header.Clone() as ToxyRow);
            }
            return newtt;
        }

        private string GetFixedLengthTableLine(int length, string value=null)
        {
            if (length <= 0)
                return string.Empty;
            StringBuilder sb = new StringBuilder();
            if (!string.IsNullOrEmpty(value))
            {
                sb.Append(' ');
                sb.Append(value);
                sb.Append(" |");
            }
            else
            {
                for (int i = 0; i < length; i++)
                {
                    sb.Append('-');
                }
                sb.Append('+');
            }

            return sb.ToString();
        }

        public string Print()
        {
            StringBuilder sb = new StringBuilder();
            if (this.HeaderRows.Count == 0)
                return null;
            
            //draw top line of the header
            sb.Append('+');
            foreach (var cell in this.HeaderRows[0].Cells)
            {
                sb.Append(GetFixedLengthTableLine(cell.Value.Length + 2));
            }

            sb.AppendLine();
            //draw the fields in the header
            sb.Append('|');
            foreach (var cell in this.HeaderRows[0].Cells)
            {
                sb.Append(GetFixedLengthTableLine(cell.Value.Length, cell.Value));
            }

            sb.AppendLine();
            //draw bottom line of the header
            sb.Append('+');
            foreach (var cell in this.HeaderRows[0].Cells)
            {
                sb.Append(GetFixedLengthTableLine(cell.Value.Length + 2));
            }
            sb.AppendLine();
            if (this.Rows.Count > 0)
            {
                foreach (var row in this.Rows)
                {
                    sb.Append('|');
                    foreach (var cell in row.Cells)
                    {
                        sb.Append(GetFixedLengthTableLine(cell.Value.Length, cell.Value));
                    }
                    sb.AppendLine();
                }
                //draw bottom line of the table
                sb.Append('+');
                foreach (var cell in this.HeaderRows[0].Cells)
                {
                    sb.Append(GetFixedLengthTableLine(cell.Value.Length + 2));
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        public string Print(int start, int end)
        {
            StringBuilder sb = new StringBuilder();
            if (this.HeaderRows.Count == 0)
                return null;

            int[] widths = new int[this.HeaderRows[0].Cells.Count];
            //draw top line of the header
            sb.Append('+');
            for (int i = start; i <= end; i++)
            {
                var cell = this.HeaderRows[0].Cells[i];
                widths[i] = cell.Value.Length+2;
                sb.Append(GetFixedLengthTableLine(widths[i]));
            }

            sb.AppendLine();
            //draw the fields in the header
            sb.Append('|');
            for (int i = start; i <= end; i++)
            {
                var cell = this.HeaderRows[0].Cells[i];
                sb.Append(GetFixedLengthTableLine(widths[i], cell.Value));
            }

            sb.AppendLine();
            //draw bottom line of the header
            sb.Append('+');
            for (int i = start; i <= end; i++)
            {
                var cell = this.HeaderRows[0].Cells[i];
                sb.Append(GetFixedLengthTableLine(widths[i]));
            }
            sb.AppendLine();
            if (this.Rows.Count > 0)
            {
                int lastRow = 0;
                foreach (var row in this.Rows)
                {
                    sb.Append('|');
                    for (int i = start; i <= end; i++)
                    {
                        var cell = row.Cells[i];
                        sb.Append(GetFixedLengthTableLine(widths[i], cell.Value));
                    }
                    sb.AppendLine();
                    lastRow++;
                }
                //draw bottom line of the table
                sb.Append('+');
                for (int i = start; i <= end; i++)
                {
                    var cell = this.Rows[lastRow-1].Cells[i];
                    sb.Append(GetFixedLengthTableLine(widths[i]));
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        public string Print(string[] fieldRange)
        {
            throw new NotImplementedException();
        }
    }
}

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

        private int[] CalculateColumnWidths(int startCol, int endCol)
        {
            int colCount = endCol - startCol + 1;
            int[] widths = new int[colCount];

            for (int i = 0; i < colCount; i++)
            {
                widths[i] = this.HeaderRows[0].Cells[startCol + i].Value.Length;
            }

            foreach (var row in this.Rows)
            {
                for (int i = 0; i < colCount; i++)
                {
                    int cellLength = row.Cells[startCol + i].Value.Length;
                    if (cellLength > widths[i])
                        widths[i] = cellLength;
                }
            }

            for (int i = 0; i < colCount; i++)
            {
                widths[i] += 2;
            }

            return widths;
        }

        private string GetTableLine(int[] widths, bool isSeparator)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append('+');
            foreach (var width in widths)
            {
                for (int i = 0; i < width; i++)
                {
                    sb.Append(isSeparator ? '-' : ' ');
                }
                sb.Append('+');
            }
            return sb.ToString();
        }

        private string FormatCell(int width, string value)
        {
            return value.PadRight(width - 1) + " |";
        }

        public void Print()
        {
            if (this.HeaderRows.Count == 0)
                return;

            int[] widths = CalculateColumnWidths(0, this.HeaderRows[0].Cells.Count - 1);
            PrintWithWidths(widths, 0, this.HeaderRows[0].Cells.Count - 1);
        }

        private void PrintWithWidths(int[] widths, int startCol, int endCol)
        {
            Console.WriteLine(GetTableLine(widths, true));
            Console.Write('|');
            for (int i = 0; i <= endCol - startCol; i++)
            {
                Console.Write(FormatCell(widths[i], this.HeaderRows[0].Cells[startCol + i].Value));
            }
            Console.WriteLine();
            Console.WriteLine(GetTableLine(widths, true));

            foreach (var row in this.Rows)
            {
                Console.Write('|');
                for (int i = 0; i <= endCol - startCol; i++)
                {
                    Console.Write(FormatCell(widths[i], row.Cells[startCol + i].Value));
                }
                Console.WriteLine();
            }

            Console.WriteLine(GetTableLine(widths, true));
        }

        public string PrintToString()
        {
            StringBuilder sb = new StringBuilder();
            if (this.HeaderRows.Count == 0)
                return null;

            int[] widths = CalculateColumnWidths(0, this.HeaderRows[0].Cells.Count - 1);
            sb.AppendLine(GetTableLine(widths, true));
            sb.Append('|');
            for (int i = 0; i < widths.Length; i++)
            {
                sb.Append(FormatCell(widths[i], this.HeaderRows[0].Cells[i].Value));
            }
            sb.AppendLine();
            sb.AppendLine(GetTableLine(widths, true));

            foreach (var row in this.Rows)
            {
                sb.Append('|');
                for (int i = 0; i < widths.Length; i++)
                {
                    sb.Append(FormatCell(widths[i], row.Cells[i].Value));
                }
                sb.AppendLine();
            }

            sb.Append(GetTableLine(widths, true));
            return sb.ToString();
        }

        public void Print(int start, int end)
        {
            if (this.HeaderRows.Count == 0)
                return;

            int[] widths = CalculateColumnWidths(start, end);
            PrintWithWidths(widths, start, end);
        }

        public string PrintToString(int start, int end)
        {
            StringBuilder sb = new StringBuilder();
            if (this.HeaderRows.Count == 0)
                return null;

            int[] widths = CalculateColumnWidths(start, end);
            sb.AppendLine(GetTableLine(widths, true));
            sb.Append('|');
            for (int i = start; i <= end; i++)
            {
                sb.Append(FormatCell(widths[i - start], this.HeaderRows[0].Cells[i].Value));
            }
            sb.AppendLine();
            sb.AppendLine(GetTableLine(widths, true));

            foreach (var row in this.Rows)
            {
                sb.Append('|');
                for (int i = start; i <= end; i++)
                {
                    sb.Append(FormatCell(widths[i - start], row.Cells[i].Value));
                }
                sb.AppendLine();
            }

            sb.Append(GetTableLine(widths, true));
            return sb.ToString();
        }

        public void Print(string[] fieldRange)
        {
            throw new NotImplementedException();
        }
    }
}

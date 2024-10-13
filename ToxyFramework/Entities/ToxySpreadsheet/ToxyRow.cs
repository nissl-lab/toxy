using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

namespace Toxy
{
    public class ToxyRow:ICloneable
    {
        public int RowIndex { get; set; }
        public int LastCellIndex { get; set; }
        public ToxyRow(int rowIndex)
        {
            this.RowIndex = rowIndex;
            this.Cells = new List<ToxyCell>();
        }

        public int Length
        {
            get { return this.Cells.Count; }
        }

        public ToxyCell this[int i]
        {
            get { 
                return this.Cells[i]; 
            }
        }

        public ToxyCell[] Slice(int start, int length)
        {
            if (length >= this.Length)
                throw new ArgumentOutOfRangeException();

            var slice = new ToxyCell[length];
            this.Cells.CopyTo(start,slice, 0, length);
            return slice;
        }

        public List<ToxyCell> Cells
        {
            get;
            set;
        }

        public object Clone()
        {
            ToxyRow clonedrow = new ToxyRow(this.RowIndex);
            foreach (ToxyCell cell in this.Cells)
            {
                clonedrow.Cells.Add(cell.Clone() as ToxyCell);
            }
            return clonedrow;
        }
    }
}

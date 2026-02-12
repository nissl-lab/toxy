using System;
using System.Collections.Generic;
using System.Linq;

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
                return this.Cells.SingleOrDefault(c=>c.CellIndex==i); 
            }
        }

        public ToxyCell[] Slice(int start, int length)
        {
            if (start<0 || this.Cells.Count==0 ||this.Cells.Max(t => t.CellIndex) < start + length - 1)
                throw new ArgumentOutOfRangeException();

            return this.Cells.Where(t => t.CellIndex >= start && t.CellIndex < start + length).ToArray();
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

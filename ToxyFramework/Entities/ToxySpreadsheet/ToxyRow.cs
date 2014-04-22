using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Toxy
{
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
}

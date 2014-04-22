using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
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
}

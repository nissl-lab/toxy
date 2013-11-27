using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IToxySpreadsheet
    {

        int TotalSheetNumber { get; set; }
        List<IProperty> Properties { get; set; }
    }
}

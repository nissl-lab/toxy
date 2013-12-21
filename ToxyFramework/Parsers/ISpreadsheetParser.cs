using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface ISpreadsheetParser
    {
        ToxySpreadsheet Parse();
        ParserContext Context { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IDomParser
    {
        ToxyDom Parse();
        ParserContext Context { get; set; }
    }
}

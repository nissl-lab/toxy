using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface ITextParser
    {
        string Parse();
        ParserContext Context { get; set; }
    }
}

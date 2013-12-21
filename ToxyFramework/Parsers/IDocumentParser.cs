using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IDocumentParser
    {
        ToxyDocument Parse();
        ParserContext Context { get; set; }
    }
}

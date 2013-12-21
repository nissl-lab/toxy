using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IEmailParser
    {
        ToxyEmail Parse();
        ParserContext Context { get; set; }
    }
}

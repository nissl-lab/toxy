using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IParser
    {
        string Parse(IParserContext context);
    }
}

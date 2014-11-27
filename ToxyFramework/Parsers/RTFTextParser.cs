using DCSoft.RTF;
using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Parsers
{
    public class RTFTextParser : PlainTextParser
    {
        public RTFTextParser(ParserContext context): base(context)
        {
            this.Context = context;
        }

        public override string Parse()
        {
            RTFDomDocument doc = new RTFDomDocument();
            doc.Load(this.Context.Path);
            return doc.InnerText;
        }
    }
}

using RtfPipe;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class RTFTextParser : PlainTextParser
    {
        public RTFTextParser(ParserContext context) : base(context)
        {
            this.Context = context;
        }

        public override string Parse()
        {
            using (var fs = new FileStream(Context.Path, FileMode.Open))
            {
                var html = Rtf.ToHtml(fs);
                return html;
            }
        }


    }
}
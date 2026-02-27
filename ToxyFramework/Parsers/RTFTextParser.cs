using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Toxy;

namespace Toxy.Parsers
{
    public class RTFTextParser : ITextParser
    {
        public RTFTextParser(ParserContext context)
        {
            this.Context = context;
        }
        public virtual ParserContext Context
        {
            get; set;
        }
        public string Parse()
        {
            Utility.ValidateContext(Context);

            byte[] bytes = null;
            if (Context.IsStreamContext)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    Context.Stream.CopyTo(ms);
                    bytes = ms.ToArray();
                }
            }
            else
            {
                bytes = File.ReadAllBytes(Context.Path);
            }
            ReasonableRTF.RtfToTextConverter converter = new ReasonableRTF.RtfToTextConverter();
            ReasonableRTF.Models.RtfResult result = converter.Convert(bytes);
            return result.Text;
        }
    }
}
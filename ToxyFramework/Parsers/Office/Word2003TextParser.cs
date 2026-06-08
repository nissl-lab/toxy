using NPOI.HWPF;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class Word2003TextParser : ITextParser
    {
        public Word2003TextParser(ParserContext context)
        {
            this.Context = context;
        }
        public string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                HWPFDocument worddoc = new HWPFDocument(stream);
                return worddoc.GetRange().Text;
            }
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

//using iTextSharp.text.pdf;
//using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class PDFTextParser : ITextParser
    {
        public PDFTextParser(ParserContext context)
        {
            this.Context = context;
        }
        public string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            return PdfTextract.PdfTextExtractor.GetText(Context.Path);
        }

        public ParserContext Context { get; set; }
    }
}
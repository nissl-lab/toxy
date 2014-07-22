using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Parsers
{
    public class PDFTextParser:ITextParser
    {
        public PDFTextParser(ParserContext context)
        {
            this.Context = context;
        }
        public string Parse()
        {
            using (PdfReader reader = new PdfReader(this.Context.Path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
                   string thePage = PdfTextExtractor.GetTextFromPage(reader, i, its);
                   text.AppendLine(thePage);
                }
                return text.ToString();
            }
        }

        public ParserContext Context { get; set; }
    }
}

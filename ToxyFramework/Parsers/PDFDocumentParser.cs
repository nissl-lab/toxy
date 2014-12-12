using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class PDFDocumentParser: IDocumentParser
    {
        public PDFDocumentParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyDocument Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            ToxyDocument rdoc = new ToxyDocument();
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();

            using (PdfReader reader = new PdfReader(this.Context.Path))
            {

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string thePage = PdfTextExtractor.GetTextFromPage(reader, i, its);
                    string[] theLines = thePage.Split('\n');
                    foreach (var theLine in theLines)
                    {
                        ToxyParagraph para = new ToxyParagraph();
                        para.Text = theLine;
                        rdoc.Paragraphs.Add(para);
                    }
                }
            }
            return rdoc;
        }
        public ParserContext Context
        {
            get;
            set;
        }
    }
}

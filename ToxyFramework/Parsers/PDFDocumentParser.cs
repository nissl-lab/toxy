using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.IO;
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
            using (Stream stream = File.OpenRead(Context.Path))

            using (PdfDocument doc = PdfReader.Open(stream, PdfDocumentOpenMode.ReadOnly))
            {
                for (int i = 0; i < doc.PageCount; i++)
                {
                    var texts = doc.Pages[i].ExtractText();
                    foreach (var text in texts)
                    {
                        ToxyParagraph para = new ToxyParagraph();
                        para.Text = text;
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

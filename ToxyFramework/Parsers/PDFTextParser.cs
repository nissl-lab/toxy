using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
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
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            using (PdfDocument reader = PdfReader.Open(this.Context.Path, PdfDocumentOpenMode.ReadOnly))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 0; i < reader.PageCount; i++)
                {
                    var lines = reader.Pages[i].ExtractWholeText();
                        text.AppendLine(lines);
                }
                return text.ToString();
            }
        }

        public ParserContext Context { get; set; }
    }
}

using System.Collections.Generic;
using System.IO;
using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace Toxy.Parsers
{
    public class PDFDocumentParser : IDocumentParser
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
            using (PdfDocument document = PdfDocument.Open(this.Context.Path))
            {
                StringBuilder text = new StringBuilder();

                for (var i = 0; i < document.NumberOfPages; i++)
                {
                    var page = document.GetPage(i + 1);
                    var words = page.GetWords();
                    var blocks = RecursiveXYCut.Instance.GetBlocks(words);

                    foreach (var block in blocks)
                    {
                        ToxyParagraph para = new ToxyParagraph();
                        para.Text = block.Text;
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

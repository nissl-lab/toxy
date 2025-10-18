using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
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

            using (PdfDocument reader = new PdfDocument(new PdfReader(this.Context.Path)))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.GetNumberOfPages(); i++)
                {
                    ITextExtractionStrategy its = new LocationTextExtractionStrategy();
                    string thePage = PdfTextExtractor.GetTextFromPage(reader.GetPage(i), its);
                    text.AppendLine(thePage);
                }
                return text.ToString();
            }
        }

        public ParserContext Context { get; set; }
    }
}
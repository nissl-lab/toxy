using PasswordProtectedChecker;
using System.IO;
using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

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

            var checker = new Checker();
            if (checker.IsFileProtected(Context.Path).Protected)
                throw new System.InvalidOperationException($"file {Context.Path} is encrypted");

            using (PdfDocument doc = PdfDocument.Open(this.Context.Path))
            {
                StringBuilder text = new StringBuilder();

                foreach (Page page in doc.GetPages())
                {
                    string pageText = ContentOrderTextExtractor.GetText(page);
                    text.AppendLine(pageText);
                }
                return text.ToString();
            }
        }

        public ParserContext Context { get; set; }
    }
}
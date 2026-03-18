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
            Utility.ValidateContext(Context);
            if (!Context.IsStreamContext)
            {
                var checker = new Checker();
                if (checker.IsFileProtected(Context.Path).Protected)
                    throw new System.InvalidOperationException($"file {Context.Path} is encrypted");
            }

            Stream stream = Utility.GetStream(Context);
            // IDK if something can throw here just to be sure
            try
            {
                using (PdfDocument doc = PdfDocument.Open(stream))
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
            finally
            {
                if (!Context.IsStreamContext)
                {
                    stream.Dispose();
                }
            }
        }

        public ParserContext Context { get; set; }
    }
}
using NPOI.XWPF.UserModel;
using PasswordProtectedChecker;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class Word2007TextParser : ITextParser
    {
		public ParserContext Context { get; set; }
		public Word2007TextParser(ParserContext context)
        {
            Context = context;
        }
        public string Parse()
        {
            Utility.ValidateContext(Context);
            Utility.ValidateContext(Context);
            bool extractHeader = false;
            if (Context.Properties.ContainsKey("ExtractHeader"))
            {
                extractHeader = Utility.IsTrue(Context.Properties["ExtractHeader"]);
            }
            bool extractFooter = false;
            if (Context.Properties.ContainsKey("ExtractFooter"))
            {
                extractFooter = Utility.IsTrue(Context.Properties["ExtractFooter"]);
            }

            StringBuilder sb = new StringBuilder();
            Stream stream = Utility.GetStream(Context);
            try
            {
                using (XWPFDocument worddoc = new XWPFDocument(stream))
                {
                    if (extractHeader && worddoc.HeaderList != null)
                    {
                        foreach (var header in worddoc.HeaderList)
                        {
                            sb.Append("[Header] ");
                            sb.AppendLine(header.Text);
                        }
                    }
                    foreach (var elem in worddoc.BodyElements)
                    {
                        if (elem is XWPFParagraph)
                        {
                            XWPFParagraph para = elem as XWPFParagraph;
                            string text = para.ParagraphText;
                            sb.AppendLine(text);
                        }
                        else if (elem is XWPFTable)
                        {
                            XWPFTable table = elem as XWPFTable;
                            string text = table.Text;
                            sb.AppendLine(text);
                        }
                    }
                    if (extractFooter && worddoc.FooterList != null)
                    {
                        foreach (var footer in worddoc.FooterList)
                        {
                            sb.Append("[Footer] ");
                            sb.AppendLine(footer.Text);
                        }
                    }
                }
            }
            finally
            {
                if (!Context.IsStreamContext)
                {
                    stream.Dispose();
                }
            }
            return sb.ToString();
        }
    }
}

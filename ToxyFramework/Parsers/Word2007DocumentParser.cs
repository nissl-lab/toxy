using NPOI.XWPF.UserModel;
using PasswordProtectedChecker;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class Word2007DocumentParser : IDocumentParser
    {
        public Word2007DocumentParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyDocument Parse()
        {
            Utility.ValidateContext(Context);

            var checker = new Checker();
            if (checker.IsFileProtected(Context.Path).Protected)
                throw new System.InvalidOperationException($"file {Context.Path} is encrypted");

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

            ToxyDocument rdoc = new ToxyDocument();

            using (var stream = Utility.GetStream(Context))
            {
                using (XWPFDocument worddoc = new XWPFDocument(stream))
                {
                    if (extractHeader && worddoc.HeaderList != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var header in worddoc.HeaderList)
                        {
                            sb.AppendLine(header.Text);
                        }
                        rdoc.Header = sb.ToString();
                    }
                    if (extractFooter && worddoc.FooterList != null)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (var footer in worddoc.FooterList)
                        {
                            sb.AppendLine(footer.Text);
                        }
                        rdoc.Footer = sb.ToString();
                    }
                    foreach (var para in worddoc.Paragraphs)
                    {
                        string text = para.ParagraphText;
                        ToxyParagraph p = new ToxyParagraph();
                        p.Text = text;
                        p.StyleID = para.Style;

                        rdoc.Paragraphs.Add(p);
                    }

                    var tables = worddoc.Tables;
                    foreach (var table in tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            var cells = row.GetTableCells();
                            foreach (var cell in cells)
                            {
                                foreach (var para in cell.Paragraphs)
                                {
                                    string text = para.ParagraphText;
                                    ToxyParagraph p = new ToxyParagraph();
                                    p.Text = text;
                                    //var runs = para.Runs;
                                    p.StyleID = para.Style;
                                    rdoc.Paragraphs.Add(p);
                                }
                            }
                        }
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

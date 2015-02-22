using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
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
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

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


            using (FileStream stream = File.OpenRead(Context.Path))
            {
                XWPFDocument worddoc = new XWPFDocument(stream);
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
                    //var runs = para.Runs;
                    p.StyleID = para.Style;

                    //for (int i = 0; i < runs.Count; i++)
                    //{
                    //    var run = runs[i];

                    //}
                    rdoc.Paragraphs.Add(p);
                }
               
                var tables = worddoc.Tables;
                foreach (var table in tables)
                {
                    foreach (var row in table.Rows)
                    {
                        var cells = row.GetTableCells();
                        foreach(var cell in cells)
                        {
                            foreach (var para in cell.Paragraphs)
                            {
                                string text = para.ParagraphText;
                                ToxyParagraph p = new ToxyParagraph();
                                p.Text = text;
                                //var runs = para.Runs;
                                p.StyleID= para.Style;
                                rdoc.Paragraphs.Add(p);
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

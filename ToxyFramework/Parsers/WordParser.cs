using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class WordParser : IDocumentParser
    {
        public WordParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyDocument Parse()
        {
            ToxyDocument rdoc = new ToxyDocument();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                XWPFDocument worddoc = new XWPFDocument(stream);
                foreach (var para in worddoc.Paragraphs)
                {
                    string text = para.ParagraphText;
                    Paragraph p = new Paragraph();
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
                        var c1 = row.GetCell(0);

                        foreach (var para in c1.Paragraphs)
                        {
                            string text = para.ParagraphText;
                            Paragraph p = new Paragraph();
                            p.Text = text;
                            //var runs = para.Runs;
                            p.StyleID= para.Style;
                            rdoc.Paragraphs.Add(p);
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

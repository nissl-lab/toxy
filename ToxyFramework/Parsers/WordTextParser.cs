using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
    public class WordTextParser:ITextParser
    {

        public string Parse()
        {
            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                XWPFDocument worddoc = new XWPFDocument(stream);
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
            }
            return sb.ToString();
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

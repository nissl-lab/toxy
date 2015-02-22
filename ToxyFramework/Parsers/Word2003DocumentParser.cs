using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class Word2003DocumentParser : IDocumentParser
    {
        public Word2003DocumentParser(ParserContext context)
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
                HWPFDocument worddoc = new HWPFDocument(stream);
                if (extractHeader && worddoc.GetHeaderStoryRange() != null)
                {
                    StringBuilder sb = new StringBuilder();
                    rdoc.Header = worddoc.GetHeaderStoryRange().Text;
                }
                if (extractFooter && worddoc.GetFootnoteRange() != null)
                {
                    StringBuilder sb = new StringBuilder();
                    rdoc.Footer = worddoc.GetFootnoteRange().Text;
                }
                for (int i=0;i<worddoc.GetRange().NumParagraphs;i++)
                {
                    Paragraph para = worddoc.GetRange().GetParagraph(i);
                    string text = para.Text;
                    ToxyParagraph p = new ToxyParagraph();
                    p.Text = text;
                    //var runs = para.Runs;
                    p.StyleID = para.GetStyleIndex().ToString();

                    //for (int i = 0; i < runs.Count; i++)
                    //{
                    //    var run = runs[i];

                    //}
                    rdoc.Paragraphs.Add(p);
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

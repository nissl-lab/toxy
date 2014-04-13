using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class Paragraph
    {
        public string Text { get; set; }
        public string StyleID { get; set; }
        public ParagraphStyle EmbededStyle { get; set; }

        public override string ToString()
        {
            return this.Text;
        }
    }
    public class ParagraphStyle
    {
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderlined { get; set; }
        public string FontFamily { get; set; }
        public int FontSize { get; set; }
    }
    public class ToxyDocument:IToxyProperties
    {
        public ToxyDocument()
        {
            this.Paragraphs = new List<Paragraph>();
            this.Properties = new Dictionary<string,object>();
        }
        public string Header { get; set; }
        public string Footer { get; set; }
        public List<Paragraph> Paragraphs { get; set; }
        public int TotalPageNumber { get; set; }
        public Dictionary<string, object> Properties { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var para in Paragraphs)
            {
                sb.AppendLine(para.Text);
            }
            return sb.ToString();
        }
    }
}

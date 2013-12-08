using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class Paragraph
    {
        public string Text { get; set; }
        public string StyleID { get; set; }
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
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyDocument
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

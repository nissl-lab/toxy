using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyEmail
    {
        public ToxyEmail()
        {
            this.From = new List<string>();
            this.To = new List<string>();
            this.Cc = new List<string>();
            this.Bcc = new List<string>();
            this.Attachments = new List<string>();
        }

        public List<string> From { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public List<string> Bcc { get; set; }
        public List<string> Attachments { get; set; }

        public string HtmlBody { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime? ArrivalTime { get; set; }
    }
}


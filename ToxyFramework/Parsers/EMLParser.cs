using HLIB.MailFormats;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class EMLParser:IEmailParser
    {
        public EMLParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyEmail Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            ToxyEmail email = new ToxyEmail();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                EMLReader reader = new EMLReader(stream);
                email.From = new List<string>(reader.From.Split(';'));
                email.To = new List<string>(reader.To.Split(';'));
                email.Cc = new List<string>(reader.CC.Split(';'));
                email.Body = reader.Body;
                email.HtmlBody = reader.HTMLBody;
                email.Subject = reader.Subject;
                email.ArrivalTime = reader.X_OriginalArrivalTime;
            }

            return email;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

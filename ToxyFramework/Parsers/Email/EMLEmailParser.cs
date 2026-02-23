using HLIB.MailFormats;
using System.Collections.Generic;
using System.IO;

namespace Toxy.Parsers
{
    public class EMLEmailParser:IEmailParser
    {
        public EMLEmailParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyEmail Parse()
        {
            if (!Context.IsStreamContext && !File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            Stream stream = null;
            if (Context.IsStreamContext)
                stream = Context.Stream;
            else
                stream = File.OpenRead(Context.Path);

            ToxyEmail email = new ToxyEmail();
            EMLReader reader = new EMLReader(stream);
            email.From = reader.From;
            email.To = new List<string>(reader.To.Split(';'));
            if (reader.CC != null)
                email.Cc = new List<string>(reader.CC.Split(';'));

            email.TextBody = reader.Body;
            email.HtmlBody = reader.HTMLBody;
            email.Subject = reader.Subject;
            email.ArrivalTime = reader.X_OriginalArrivalTime;
            return email;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

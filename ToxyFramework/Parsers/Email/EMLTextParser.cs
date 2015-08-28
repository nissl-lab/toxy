using HLIB.MailFormats;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class EMLTextParser:PlainTextParser
    {
        public EMLTextParser(ParserContext context):base(context)
        {
            this.Context = context;
        }

        public override string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            StringBuilder sb = new StringBuilder();
            using (FileStream stream = File.OpenRead(Context.Path))
            {
                EMLReader reader = new EMLReader(stream);
                if (!string.IsNullOrEmpty(reader.From))
                    sb.AppendFormat("[From] {0}{1}",reader.From, Environment.NewLine);
                if (!string.IsNullOrEmpty(reader.To))
                    sb.AppendFormat("[To] {0}{1}", reader.To, Environment.NewLine);
                if (!string.IsNullOrEmpty(reader.CC))
                    sb.AppendFormat("[CC] {0}{1}", reader.CC, Environment.NewLine);
                if (!string.IsNullOrEmpty(reader.Subject))
                    sb.AppendFormat("[Subject] {0}{1}", reader.Subject, Environment.NewLine);

                sb.AppendLine();
                sb.AppendLine(reader.Body);
                //sb.AppendLine(reader.HTMLBody);
            }
            return sb.ToString();
        }
    }
}

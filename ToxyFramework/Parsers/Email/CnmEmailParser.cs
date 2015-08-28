using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Toxy.Parsers
{
    public class CnmEmailParser : IEmailParser
    {
        public CnmEmailParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxyEmail Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            FileInfo fi = new FileInfo(Context.Path);
            CnmFile reader = new CnmFile(fi);
            reader.Parse();
            ToxyEmail email=new ToxyEmail();
            email.From = reader.From;
            email.To = new List<string>(reader.To.Split(';'));
            email.TextBody = reader.TextPlain;
            using (var sr=new StreamReader(reader.TextHtml))
            {
                email.HtmlBody = sr.ReadToEnd();
            }
            email.Subject = reader.Subject;
            return email;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

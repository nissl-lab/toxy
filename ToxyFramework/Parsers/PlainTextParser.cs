using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class PlainTextParser:ITextParser
    {
        public class ParseLineEventArgs : EventArgs
        {
            public ParseLineEventArgs(string text, int number)
            {
                this.LineNumber = number;
                this.Text = text;
            }
            public int LineNumber { get; set; }
            public string Text { get; set; }
        }


        public event EventHandler<ParseLineEventArgs> ParseLine;

        public string Parse(ParserContext context)
        {
            if (!File.Exists(context.Path))
                throw new FileNotFoundException("File "+context.Path+" is not found");

            Encoding encoding= Encoding.UTF8;
            StreamReader sr = null;
            try
            {
                if (context.Encoding == null)
                {
                    sr = new StreamReader(context.Path, true);
                }
                else
                {
                    sr = new StreamReader(context.Path, true);
                }
                string line = sr.ReadLine();
                int i = 0;
                StringBuilder sb = new StringBuilder();
                while(line!=null)
                {
                    if (ParseLine != null)
                    {
                        ParseLine(this, new ParseLineEventArgs(line, i));
                    }
                    sb.AppendLine(line);
                    line = sr.ReadLine();
                    i++;
                }
                return sb.ToString();
            }
            finally
            {
                if (sr != null)
                    sr.Close();
            }
        }
    }
}

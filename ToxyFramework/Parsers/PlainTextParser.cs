using System;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class PlainTextParser : ITextParser
    {
        public class ParseLineEventArgs : EventArgs
        {
            public ParseLineEventArgs(string text, int number)
            {
                LineNumber = number;
                Text = text;
            }
            public int LineNumber { get; set; }
            public string Text { get; set; }
        }
		public ParserContext Context { get; set; }
		public PlainTextParser(ParserContext context)
        {
            Context = context;
        }
        public event EventHandler<ParseLineEventArgs> ParseLine;

        public virtual string Parse()
        {
            Utility.ValidateContext(Context);

            StreamReader sr = null;
            try
            {
                if (Context.Encoding == null)
                {
                    sr = new StreamReader(Context.Path, true);
                }
                else
                {
                    sr = new StreamReader(Context.Path, Context.Encoding);
                }
                string line = sr.ReadLine();
                int i = 0;
                StringBuilder sb = new StringBuilder();
                while (line != null)
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

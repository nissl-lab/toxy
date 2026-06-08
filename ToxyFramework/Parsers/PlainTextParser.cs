using System;
using System.IO;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	public class PlainTextParser : BaseTextParser
	{
		public event EventHandler<ParseLineEventArgs> ParseLine;
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

		public PlainTextParser(ParserContext context) : base(context) { }

		internal override string ParseText(ref IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			Encoding encoding = Context.Encoding ?? Encoding.UTF8;
			// this is what new StreamReader(path, Context.Encoding) & new StreamReader(path, true) does
			StreamReader sr = new StreamReader(stream, encoding, true, 1024, !Context.IsStreamContext);
			disposable = sr;

			string line = sr.ReadLine();
			int i = 0;
			StringBuilder sb = new StringBuilder();
			while (line != null)
			{
				ParseLine?.Invoke(this, new ParseLineEventArgs(line, i));
				sb.AppendLine(line);
				line = sr.ReadLine();
				i++;
			}
			return sb.ToString();
		}
	}
}

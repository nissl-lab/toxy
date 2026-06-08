using HtmlAgilityPack;
using System;
using System.Text;
using System.Text.RegularExpressions;
using Toxy.Base;
using VersOne.Epub;

namespace Toxy.Parsers
{
	public sealed class EPUBTextParser : BaseTextParser
	{
		public EPUBTextParser(ParserContext context) : base(context) { }

		internal override string ParseText(ref IDisposable disposable)
		{
			EpubBook book = EPUBHelper.GetEpubBook(Context);
			StringBuilder result = new StringBuilder();
			HtmlDocument html = new HtmlDocument();
			Regex whiteSpaceStart = new Regex("^\\s*", RegexOptions.Multiline);
			foreach (EpubLocalTextContentFile chapter in book.ReadingOrder)
			{
				html.LoadHtml(chapter.Content);
				result.AppendLine(whiteSpaceStart.Replace(html.DocumentNode.InnerText, ""));
			}
			return result.ToString();
		}
	}
}

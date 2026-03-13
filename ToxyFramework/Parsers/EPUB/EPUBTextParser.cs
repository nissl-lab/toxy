using HtmlAgilityPack;
using System.Text;
using System.Text.RegularExpressions;
using VersOne.Epub;

namespace Toxy.Parsers
{
	public class EPUBTextParser : ITextParser
	{
		public EPUBTextParser(ParserContext context)
		{
			Context = context;
		}
		public virtual ParserContext Context { get; set; }
		public string Parse()
		{
			EpubBook book = EPUBHelper.GetEpubBook(Context);
			StringBuilder result = new StringBuilder();
			HtmlDocument html = new HtmlDocument();
			Regex whiteSpaceStart = new Regex("^\\s*", RegexOptions.Multiline);
			foreach (EpubLocalTextContentFile chapter in book.ReadingOrder)
			{
				html.LoadHtml(chapter.Content);
				var a = whiteSpaceStart.Replace(html.DocumentNode.InnerText, "");
				result.AppendLine(whiteSpaceStart.Replace(html.DocumentNode.InnerText, ""));
			}
			return result.ToString();
		}
	}
}

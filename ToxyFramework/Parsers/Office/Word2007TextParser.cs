using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	public class Word2007TextParser : BaseTextParser
	{
		public Word2007TextParser(ParserContext context) : base(context)
		{ }

		internal override void ValidateContext()
		{
			base.ValidateContext();
			Utility.ThrowIfProtected(Context);
		}

		internal override string ParseText(Stream stream)
		{
			bool extractHeader = false;
			if (Context.Properties.ContainsKey("ExtractHeader"))
			{
				extractHeader = Utility.IsTrue(Context.Properties["ExtractHeader"]);
			}
			bool extractFooter = false;
			if (Context.Properties.ContainsKey("ExtractFooter"))
			{
				extractFooter = Utility.IsTrue(Context.Properties["ExtractFooter"]);
			}

			StringBuilder sb = new StringBuilder();
			using (WordprocessingDocument worddoc = WordprocessingDocument.Open(stream, false))
			{
				MainDocumentPart mainPart = worddoc.MainDocumentPart;
				if (extractHeader && mainPart != null)
				{
					foreach (HeaderPart headerPart in mainPart.HeaderParts)
					{
						sb.Append("[Header] ");
						sb.AppendLine(WordprocessingText.GetText(headerPart.Header));
					}
				}

				Body body = mainPart?.Document?.Body;
				if (body != null)
				{
					foreach (var block in WordprocessingText.GetBlocks(body))
					{
						if (block is Paragraph paragraph)
						{
							sb.AppendLine(WordprocessingText.GetText(paragraph));
						}
						else if (block is Table table)
						{
							sb.AppendLine(WordprocessingText.GetText(table));
						}
					}
				}

				if (extractFooter && mainPart != null)
				{
					foreach (FooterPart footerPart in mainPart.FooterParts)
					{
						sb.Append("[Footer] ");
						sb.AppendLine(WordprocessingText.GetText(footerPart.Footer));
					}
				}
			}
			return sb.ToString();
		}
	}
}

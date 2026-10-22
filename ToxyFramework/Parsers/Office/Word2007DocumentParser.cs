using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
	public class Word2007DocumentParser : IDocumentParser
	{
		public ParserContext Context { get; set; }

		public Word2007DocumentParser(ParserContext context)
		{
			Context = context;
		}

		public ToxyDocument Parse()
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);
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

			ToxyDocument rdoc = new ToxyDocument();
			Stream stream = Utility.GetStream(Context);
			try
			{
				using (WordprocessingDocument worddoc = WordprocessingDocument.Open(stream, false))
				{
					MainDocumentPart mainPart = worddoc.MainDocumentPart;
					if (mainPart == null)
					{
						return rdoc;
					}

					if (extractHeader)
					{
						StringBuilder sb = new StringBuilder();
						foreach (HeaderPart headerPart in mainPart.HeaderParts)
						{
							sb.AppendLine(WordprocessingText.GetText(headerPart.Header));
						}
						rdoc.Header = sb.ToString();
					}
					if (extractFooter)
					{
						StringBuilder sb = new StringBuilder();
						foreach (FooterPart footerPart in mainPart.FooterParts)
						{
							sb.AppendLine(WordprocessingText.GetText(footerPart.Footer));
						}
						rdoc.Footer = sb.ToString();
					}

					Body body = mainPart.Document?.Body;
					if (body != null)
					{
						// body paragraphs first, followed by the paragraphs of the tables
						foreach (Paragraph para in WordprocessingText.GetParagraphs(body))
						{
							rdoc.Paragraphs.Add(ToToxyParagraph(para));
						}
						foreach (Table table in WordprocessingText.GetTables(body))
						{
							foreach (Paragraph para in WordprocessingText.GetCellParagraphs(table))
							{
								rdoc.Paragraphs.Add(ToToxyParagraph(para));
							}
						}
					}
				}
			}
			finally
			{
				if (!Context.IsStreamContext)
				{
					stream.Dispose();
				}
			}
			return rdoc;
		}

		private static ToxyParagraph ToToxyParagraph(Paragraph paragraph)
		{
			ToxyParagraph p = new ToxyParagraph();
			p.Text = WordprocessingText.GetText(paragraph);
			p.StyleID = WordprocessingText.GetStyleId(paragraph);
			return p;
		}
	}
}

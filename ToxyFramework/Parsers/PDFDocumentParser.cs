using PasswordProtectedChecker;
using System.Collections.Generic;
using System.IO;
using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

namespace Toxy.Parsers
{
	public class PDFDocumentParser : IDocumentParser
	{
		public ParserContext Context { get; set; }
		public PDFDocumentParser(ParserContext context)
		{
			Context = context;
		}

		public ToxyDocument Parse()
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);

			ToxyDocument rdoc = new ToxyDocument();
			Stream stream = Utility.GetStream(Context);
			// IDK if something throws there
			try
			{
				using (PdfDocument document = PdfDocument.Open(stream))
				{
					StringBuilder text = new StringBuilder();

					for (var i = 0; i < document.NumberOfPages; i++)
					{
						var page = document.GetPage(i + 1);
						var words = page.GetWords();
						var blocks = RecursiveXYCut.Instance.GetBlocks(words);

						foreach (var block in blocks)
						{
							ToxyParagraph para = new ToxyParagraph();
							para.Text = block.Text;
							rdoc.Paragraphs.Add(para);
						}
					}
				}
			}
			finally
			{
				if (Context.IsStreamContext)
				{
					stream.Dispose();
				}
			}
			return rdoc;
		}

	}
}

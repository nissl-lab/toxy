using PasswordProtectedChecker;
using System;
using System.IO;
using System.Text;
using Toxy.Base;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Toxy.Parsers
{
	public class PDFTextParser : BaseTextParser
	{
		public PDFTextParser(ParserContext context) : base(context)
		{ }

		internal override void ValidateContext()
		{
			base.ValidateContext();
			Utility.ThrowIfProtected(Context);
		}

		internal override string ParseText(Stream stream)
		{
			using (PdfDocument doc = PdfDocument.Open(stream))
			{
				StringBuilder text = new StringBuilder();

				foreach (Page page in doc.GetPages())
				{
					string pageText = ContentOrderTextExtractor.GetText(page);
					text.AppendLine(pageText);
				}
				return text.ToString();
			}
		}
	}
}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PasswordProtectedChecker;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	public class Powerpoint2007TextParser : BaseTextParser
	{
		public Powerpoint2007TextParser(ParserContext context) : base(context)
		{ }

		internal override void ValidateContext()
		{
			base.ValidateContext();
			Utility.ThrowIfProtected(Context);
		}

		internal override string ParseText(Stream stream)
		{
			StringBuilder sb = new StringBuilder();
			// closes the Stream anyways
			using (Package pkg = Package.Open(stream, FileMode.Open, FileAccess.Read))
			{
				using PresentationDocument ppt = PresentationDocument.Open(pkg);
				// Get the relationship ID of the first slide.
				PresentationPart part = ppt.PresentationPart;
				OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
				for (int index = 0; index < slideIds.Count; index++)
				{
					string relId = (slideIds[index] as SlideId).RelationshipId;

					// Get the slide part from the relationship ID.
					SlidePart slide = (SlidePart)part.GetPartById(relId);
					string[] texts = GetAllTextInSlide(slide);
					foreach (string text in texts)
					{
						sb.AppendLine(text);
					}
				}
			}
			return sb.ToString();
		}

		public static string[] GetAllTextInSlide(SlidePart slidePart)
		{
			// Verify that the slide part exists.
#if NET8_0_OR_GREATER
			ArgumentNullException.ThrowIfNull(slidePart, nameof(slidePart));
#else
			if (slidePart == null)
			{
				throw new ArgumentNullException("slidePart");
			}
#endif

			// Create a new linked list of strings.
			LinkedList<string> texts = new LinkedList<string>();

			// If the slide exists...
			if (slidePart.Slide != null)
			{
				// Iterate through all the paragraphs in the slide.
				foreach (DocumentFormat.OpenXml.Drawing.Paragraph paragraph in
					slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
				{
					// Create a new string builder.                    
					StringBuilder paragraphText = new StringBuilder();

					// Iterate through the lines of the paragraph.
					foreach (DocumentFormat.OpenXml.Drawing.Text text in
						paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
					{
						// Append each line to the previous lines.
						paragraphText.Append(text.Text);
					}

					if (paragraphText.Length > 0)
					{
						// Add each paragraph to the linked list.
						texts.AddLast(paragraphText.ToString());
					}
				}
			}

			if (texts.Count > 0)
			{
				// Return an array of strings.
				return texts.ToArray();
			}
			else
			{
				return null;
			}
		}
	}
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
	/// <summary>
	/// Text extraction on top of the Open XML SDK model of a .docx file.
	/// </summary>
	internal static class WordprocessingText
	{
		/// <summary>
		/// Enumerates the paragraphs and tables directly contained in <paramref name="container"/> in document order,
		/// including the ones wrapped in content controls.
		/// </summary>
		public static IEnumerable<OpenXmlElement> GetBlocks(OpenXmlElement container)
		{
			foreach (OpenXmlElement child in container.ChildElements)
			{
				switch (child)
				{
					case Paragraph _:
					case Table _:
						yield return child;
						break;
					case SdtBlock sdt:
						SdtContentBlock content = sdt.SdtContentBlock;
						if (content != null)
						{
							foreach (OpenXmlElement block in GetBlocks(content))
							{
								yield return block;
							}
						}
						break;
					case CustomXmlBlock custom:
						foreach (OpenXmlElement block in GetBlocks(custom))
						{
							yield return block;
						}
						break;
				}
			}
		}

		public static IEnumerable<Paragraph> GetParagraphs(OpenXmlElement container)
		{
			return GetBlocks(container).OfType<Paragraph>();
		}

		public static IEnumerable<Table> GetTables(OpenXmlElement container)
		{
			return GetBlocks(container).OfType<Table>();
		}

		/// <summary>The paragraphs directly contained in the cells of <paramref name="table"/>.</summary>
		public static IEnumerable<Paragraph> GetCellParagraphs(Table table)
		{
			return table.Elements<TableRow>()
				.SelectMany(row => row.Elements<TableCell>())
				.SelectMany(cell => GetParagraphs(cell));
		}

		public static string GetStyleId(Paragraph paragraph)
		{
			return paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
		}

		/// <summary>The text of the runs of a paragraph. Tabs and line, page and column breaks are rendered as '\t' and '\n'.</summary>
		public static string GetText(Paragraph paragraph)
		{
			StringBuilder builder = new StringBuilder();
			foreach (Run run in paragraph.Descendants<Run>())
			{
				// runs of paragraphs nested in the paragraph (e.g. text boxes) belong to the nested paragraph
				if (run.Ancestors<Paragraph>().First() != paragraph)
				{
					continue;
				}
				foreach (OpenXmlElement element in run.ChildElements)
				{
					switch (element)
					{
						case Text text: builder.Append(text.Text); break;
						case TabChar _: builder.Append('\t'); break;
						case Break _: builder.Append('\n'); break;
						case CarriageReturn _: builder.Append('\n'); break;
						case NoBreakHyphen _: builder.Append('-'); break;
					}
				}
			}
			return builder.ToString();
		}

		/// <summary>The text of a table: one line per row, cells separated by tabs.</summary>
		public static string GetText(Table table)
		{
			StringBuilder builder = new StringBuilder();
			foreach (TableRow row in table.Elements<TableRow>())
			{
				IEnumerable<string> cells = row.Elements<TableCell>()
					.Select(cell => string.Join("\n", GetParagraphs(cell).Select(GetText)));
				builder.Append(string.Join("\t", cells)).Append('\n');
			}
			return builder.ToString();
		}

		/// <summary>The text of a header, footer or other container of block elements, one line per paragraph.</summary>
		public static string GetText(OpenXmlElement container)
		{
			IEnumerable<string> blocks = GetBlocks(container).Select(block =>
				block is Paragraph paragraph ? GetText(paragraph) : GetText((Table)block));
			return string.Join("\n", blocks);
		}
	}
}

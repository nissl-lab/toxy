using System.Collections.Generic;
using System.Text;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Renders a <see cref="WorksheetData"/> as plain text: one line per row, cells separated by tabs.
	/// </summary>
	internal static class ExcelTextWriter
	{
		public static void AppendSheet(StringBuilder builder, WorksheetData sheet, ExcelParseOptions options)
		{
			if (options.IncludeSheetNames)
			{
				builder.Append(sheet.Name).Append('\n');
			}
			if (options.IncludeHeaderFooter)
			{
				AppendHeaderFooter(builder, sheet.Header);
			}

			foreach (RowData row in sheet.Rows)
			{
				var texts = new List<string>(row.Cells.Count);
				foreach (CellData cell in row.Cells)
				{
					string text = GetText(cell, options);
					if (text.Length > 0)
					{
						texts.Add(text);
					}
				}
				if (texts.Count > 0)
				{
					builder.Append(string.Join("\t", texts)).Append('\n');
				}
			}

			if (options.IncludeHeaderFooter)
			{
				AppendHeaderFooter(builder, sheet.Footer);
			}
		}

		private static string GetText(CellData cell, ExcelParseOptions options)
		{
			string text = !options.ShowCalculatedResult && cell.Formula != null ? cell.Formula : cell.Text;
			if (options.IncludeComments && !string.IsNullOrEmpty(cell.Comment))
			{
				text += " Comment by " + cell.CommentAuthor + ": " + cell.Comment.Replace('\n', ' ');
			}
			return text;
		}

		private static void AppendHeaderFooter(StringBuilder builder, HeaderFooterText text)
		{
			var parts = new List<string>(3);
			foreach (string part in new[] { text.Left, text.Center, text.Right })
			{
				if (part.Length > 0)
				{
					parts.Add(part);
				}
			}
			if (parts.Count > 0)
			{
				builder.Append(string.Join("\t", parts)).Append('\n');
			}
		}
	}
}

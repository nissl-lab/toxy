using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Reads Office Open XML workbooks (.xlsx) with the Open XML SDK.
	/// </summary>
	/// <remarks>
	/// The sheet data is read row by row with an <see cref="OpenXmlReader"/>, a worksheet is never loaded as a whole.
	/// Rows and cells are reported as they are stored in the file, which includes rows of formatted blank cells.
	/// Only the first cell of a shared formula stores the formula text, the other cells of the group report
	/// <c>null</c> as <see cref="CellData.Formula"/> and so fall back to their calculated value.
	/// </remarks>
	internal sealed class OpenXmlWorkbookReader : IWorkbookReader
	{
		private readonly Stream _stream;
		private readonly bool _ownsStream;
		private readonly SpreadsheetDocument _document;
		private readonly List<KeyValuePair<string, WorksheetPart>> _sheets = new List<KeyValuePair<string, WorksheetPart>>();
		private readonly List<string> _sharedStrings = new List<string>();
		// number format code of each cell format (cellXfs), null means "General"
		private readonly List<string> _numberFormats = new List<string>();
		private readonly bool _date1904;

		public OpenXmlWorkbookReader(Stream stream, bool ownsStream)
		{
			_stream = stream;
			_ownsStream = ownsStream;
			_document = SpreadsheetDocument.Open(stream, false);
			try
			{
				WorkbookPart workbookPart = _document.WorkbookPart;
				if (workbookPart?.Workbook == null)
				{
					throw new InvalidDataException("The document does not contain a workbook.");
				}
				_date1904 = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value ?? false;

				if (workbookPart.Workbook.Sheets != null)
				{
					foreach (Sheet sheet in workbookPart.Workbook.Sheets.Elements<Sheet>())
					{
						// chart sheets have no cells and are skipped
						if (sheet.Id?.Value != null
							&& workbookPart.TryGetPartById(sheet.Id.Value, out OpenXmlPart part)
							&& part is WorksheetPart worksheetPart)
						{
							_sheets.Add(new KeyValuePair<string, WorksheetPart>(sheet.Name?.Value ?? string.Empty, worksheetPart));
						}
					}
				}
				ReadSharedStrings(workbookPart.SharedStringTablePart);
				ReadNumberFormats(workbookPart.WorkbookStylesPart?.Stylesheet);
			}
			catch
			{
				_document.Dispose();
				throw;
			}
		}

		public int SheetCount => _sheets.Count;

		public WorksheetData ReadSheet(int index)
		{
			if (index < 0 || index >= _sheets.Count)
			{
				throw new ArgumentOutOfRangeException(nameof(index));
			}
			WorksheetPart part = _sheets[index].Value;
			Dictionary<CellAddress, CommentData> comments = ReadComments(part.WorksheetCommentsPart);

			var cellsByRow = new SortedDictionary<int, List<CellData>>();
			var mergedRegions = new List<MergedRegion>();
			HeaderFooter headerFooter = null;
			int columnCount = 0;
			int rowIndex = -1;

			foreach (OpenXmlElement element in ReadElements(part, typeof(Row), typeof(MergeCell), typeof(HeaderFooter)))
			{
				switch (element)
				{
					case Row row:
						// the row and cell references are optional, a missing one means "next to the previous one"
						rowIndex = row.RowIndex?.Value != null ? (int)row.RowIndex.Value - 1 : rowIndex + 1;
						List<CellData> cells = GetRowCells(cellsByRow, rowIndex);
						int column = -1;
						foreach (Cell cell in row.Elements<Cell>())
						{
							column = TryParseReference(cell.CellReference?.Value, out CellAddress address) ? address.Column : column + 1;
							columnCount = Math.Max(columnCount, column + 1);

							var key = new CellAddress(rowIndex, column);
							comments.TryGetValue(key, out CommentData comment);
							comments.Remove(key);
							CellData data = ReadCell(cell, column, comment);
							if (data != null)
							{
								cells.Add(data);
							}
						}
						break;
					case MergeCell mergeCell:
						string[] range = (mergeCell.Reference?.Value ?? string.Empty).Split(':');
						if (TryParseReference(range[0], out CellAddress first)
							&& TryParseReference(range[range.Length - 1], out CellAddress last))
						{
							mergedRegions.Add(new MergedRegion(first.Row, first.Column, last.Row, last.Column));
						}
						break;
					case HeaderFooter hf:
						headerFooter = hf;
						break;
				}
			}

			// comments of cells which are not stored in the sheet data
			foreach (KeyValuePair<CellAddress, CommentData> comment in comments)
			{
				GetRowCells(cellsByRow, comment.Key.Row).Add(
					new CellData(comment.Key.Column, null, null, comment.Value.Text, comment.Value.Author));
				columnCount = Math.Max(columnCount, comment.Key.Column + 1);
			}

			var rows = new List<RowData>(cellsByRow.Count);
			foreach (KeyValuePair<int, List<CellData>> row in cellsByRow)
			{
				rows.Add(new RowData(row.Key, row.Value));
			}

			return new WorksheetData(
				index,
				_sheets[index].Key,
				ReadHeaderFooter(headerFooter?.OddHeader?.Text, headerFooter?.FirstHeader?.Text),
				ReadHeaderFooter(headerFooter?.OddFooter?.Text, headerFooter?.FirstFooter?.Text),
				columnCount,
				rows,
				mergedRegions);
		}

		public void Dispose()
		{
			_document.Dispose();
			if (_ownsStream)
			{
				_stream.Dispose();
			}
		}

		private CellData ReadCell(Cell cell, int column, CommentData comment)
		{
			string formula = cell.CellFormula?.Text;
			if (string.IsNullOrEmpty(formula))
			{
				formula = null;
			}

			object value = ReadValue(cell);
			if (value == null && comment == null && formula == null)
			{
				return null;
			}

			string text = CellValueFormatter.Format(value, GetNumberFormat(cell), _date1904);
			return new CellData(column, text, formula, comment?.Text, comment?.Author);
		}

		private object ReadValue(Cell cell)
		{
			CellValues type = cell.DataType?.Value ?? CellValues.Number;
			if (type == CellValues.InlineString)
			{
				return cell.InlineString != null ? GetText(cell.InlineString) : null;
			}

			string raw = cell.CellValue?.Text;
			if (raw == null)
			{
				return null;
			}
			if (type == CellValues.SharedString)
			{
				return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int i) && i >= 0 && i < _sharedStrings.Count
					? _sharedStrings[i]
					: null;
			}
			if (type == CellValues.Boolean)
			{
				return raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
			}
			if (type == CellValues.Date)
			{
				return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime date) ? date : (object)raw;
			}
			if (type == CellValues.String || type == CellValues.Error)
			{
				return raw;
			}
			return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : (object)raw;
		}

		private string GetNumberFormat(Cell cell)
		{
			int style = (int)(cell.StyleIndex?.Value ?? 0);
			return style < _numberFormats.Count ? _numberFormats[style] : null;
		}

		private void ReadSharedStrings(SharedStringTablePart part)
		{
			if (part == null)
			{
				return;
			}
			foreach (OpenXmlElement item in ReadElements(part, typeof(SharedStringItem)))
			{
				_sharedStrings.Add(GetText(item));
			}
		}

		private void ReadNumberFormats(Stylesheet stylesheet)
		{
			if (stylesheet?.CellFormats == null)
			{
				return;
			}
			var customFormats = new Dictionary<uint, string>();
			if (stylesheet.NumberingFormats != null)
			{
				foreach (NumberingFormat format in stylesheet.NumberingFormats.Elements<NumberingFormat>())
				{
					if (format.NumberFormatId?.Value != null)
					{
						customFormats[format.NumberFormatId.Value] = format.FormatCode?.Value;
					}
				}
			}
			foreach (CellFormat cellFormat in stylesheet.CellFormats.Elements<CellFormat>())
			{
				uint id = cellFormat.NumberFormatId?.Value ?? 0;
				_numberFormats.Add(customFormats.TryGetValue(id, out string code) ? code : CellValueFormatter.GetBuiltInFormat((int)id));
			}
		}

		private static Dictionary<CellAddress, CommentData> ReadComments(WorksheetCommentsPart part)
		{
			var result = new Dictionary<CellAddress, CommentData>();
			Comments comments = part?.Comments;
			if (comments?.CommentList == null)
			{
				return result;
			}
			var authors = new List<string>();
			if (comments.Authors != null)
			{
				foreach (Author author in comments.Authors.Elements<Author>())
				{
					authors.Add(author.Text);
				}
			}
			foreach (Comment comment in comments.CommentList.Elements<Comment>())
			{
				if (TryParseReference(comment.Reference?.Value, out CellAddress address))
				{
					int authorId = (int)(comment.AuthorId?.Value ?? 0);
					string author = authorId < authors.Count ? authors[authorId] : null;
					result[address] = new CommentData(GetText(comment.CommentText), author);
				}
			}
			return result;
		}

		private static HeaderFooterText ReadHeaderFooter(string odd, string first)
		{
			HeaderFooterText oddText = HeaderFooterTextParser.Parse(odd);
			HeaderFooterText firstText = HeaderFooterTextParser.Parse(first);
			return new HeaderFooterText(
				oddText.Left.Length > 0 ? oddText.Left : firstText.Left,
				oddText.Center.Length > 0 ? oddText.Center : firstText.Center,
				oddText.Right.Length > 0 ? oddText.Right : firstText.Right);
		}

		private static List<CellData> GetRowCells(SortedDictionary<int, List<CellData>> cellsByRow, int rowIndex)
		{
			if (!cellsByRow.TryGetValue(rowIndex, out List<CellData> cells))
			{
				cells = new List<CellData>();
				cellsByRow.Add(rowIndex, cells);
			}
			return cells;
		}

		/// <summary>
		/// Gets the text of a rich text element (shared string, inline string, comment), phonetic runs are left out.
		/// </summary>
		private static string GetText(OpenXmlElement element)
		{
			if (element == null)
			{
				return string.Empty;
			}
			var builder = new StringBuilder();
			foreach (OpenXmlElement child in element.ChildElements)
			{
				if (child is Text text)
				{
					builder.Append(text.Text);
				}
				else if (child is Run run && run.Text != null)
				{
					builder.Append(run.Text.Text);
				}
			}
			return builder.ToString();
		}

		/// <summary>
		/// Streams the elements of the given types out of <paramref name="part"/>, each one loaded with its children.
		/// </summary>
		private static IEnumerable<OpenXmlElement> ReadElements(OpenXmlPart part, params Type[] elementTypes)
		{
			using (OpenXmlReader reader = OpenXmlReader.Create(part))
			{
				bool positioned = reader.Read();
				while (positioned && !reader.EOF)
				{
					// LoadCurrentElement already moves the reader to the next element
					if (reader.IsStartElement && Array.IndexOf(elementTypes, reader.ElementType) >= 0)
					{
						yield return reader.LoadCurrentElement();
					}
					else
					{
						positioned = reader.Read();
					}
				}
			}
		}

		/// <summary>
		/// Parses an A1 style cell reference (e.g. <c>B12</c> or <c>$B$12</c>) into a zero based address.
		/// </summary>
		private static bool TryParseReference(string reference, out CellAddress address)
		{
			address = default(CellAddress);
			if (string.IsNullOrEmpty(reference))
			{
				return false;
			}
			int column = 0;
			int row = 0;
			int i = 0;
			for (; i < reference.Length && (char.IsLetter(reference[i]) || reference[i] == '$'); i++)
			{
				if (reference[i] != '$')
				{
					column = column * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
				}
			}
			for (; i < reference.Length && (char.IsDigit(reference[i]) || reference[i] == '$'); i++)
			{
				if (reference[i] != '$')
				{
					row = row * 10 + (reference[i] - '0');
				}
			}
			if (i != reference.Length || column == 0 || row == 0)
			{
				return false;
			}
			address = new CellAddress(row - 1, column - 1);
			return true;
		}

		private readonly struct CellAddress : IEquatable<CellAddress>
		{
			public CellAddress(int row, int column)
			{
				Row = row;
				Column = column;
			}

			public int Row { get; }
			public int Column { get; }

			public bool Equals(CellAddress other) => Row == other.Row && Column == other.Column;
			public override bool Equals(object obj) => obj is CellAddress other && Equals(other);
			public override int GetHashCode() => Row * 16384 + Column; // a sheet has at most 16384 columns
		}

		private sealed class CommentData
		{
			public CommentData(string text, string author)
			{
				Text = text;
				Author = author;
			}

			public string Text { get; }
			public string Author { get; }
		}
	}
}

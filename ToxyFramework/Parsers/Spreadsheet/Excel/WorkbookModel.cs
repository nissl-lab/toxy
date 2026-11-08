using System;
using System.Collections.Generic;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Text of one header or footer of a worksheet, split into its three sections.
	/// </summary>
	internal sealed class HeaderFooterText
	{
		public static readonly HeaderFooterText Empty = new HeaderFooterText(string.Empty, string.Empty, string.Empty);

		public HeaderFooterText(string left, string center, string right)
		{
			Left = left ?? string.Empty;
			Center = center ?? string.Empty;
			Right = right ?? string.Empty;
		}

		public string Left { get; }
		public string Center { get; }
		public string Right { get; }
	}

	/// <summary>
	/// A cell as read from a workbook, independent of the library used for reading it.
	/// </summary>
	internal sealed class CellData
	{
		public CellData(int columnIndex, string text, string formula, string comment, string commentAuthor)
		{
			ColumnIndex = columnIndex;
			Text = text ?? string.Empty;
			Formula = formula;
			Comment = comment;
			CommentAuthor = commentAuthor;
		}

		/// <summary>Zero based column index.</summary>
		public int ColumnIndex { get; }

		/// <summary>The displayed value, formatted according to the number format of the cell.</summary>
		public string Text { get; }

		/// <summary>The formula without leading '=', or <c>null</c> if the cell contains no formula.</summary>
		public string Formula { get; }

		/// <summary>The text of the cell comment, or <c>null</c> if there is none.</summary>
		public string Comment { get; }

		public string CommentAuthor { get; }
	}

	internal sealed class RowData
	{
		private readonly Dictionary<int, CellData> _cellsByColumn = new Dictionary<int, CellData>();

		public RowData(int rowIndex, IEnumerable<CellData> cells)
		{
			RowIndex = rowIndex;
			var ordered = new List<CellData>(cells);
			ordered.Sort((a, b) => a.ColumnIndex.CompareTo(b.ColumnIndex));
			Cells = ordered;
			foreach (CellData cell in ordered)
			{
				_cellsByColumn[cell.ColumnIndex] = cell;
			}
		}

		/// <summary>Zero based row index.</summary>
		public int RowIndex { get; }

		/// <summary>The cells of this row which carry a value, a formula or a comment, ordered by column.</summary>
		public IReadOnlyList<CellData> Cells { get; }

		/// <summary>Gets the cell at <paramref name="columnIndex"/> or <c>null</c> if the row has no such cell.</summary>
		public CellData GetCell(int columnIndex)
		{
			_cellsByColumn.TryGetValue(columnIndex, out CellData cell);
			return cell;
		}
	}

	internal sealed class MergedRegion
	{
		public MergedRegion(int firstRow, int firstColumn, int lastRow, int lastColumn)
		{
			FirstRow = firstRow;
			FirstColumn = firstColumn;
			LastRow = lastRow;
			LastColumn = lastColumn;
		}

		public int FirstRow { get; }
		public int FirstColumn { get; }
		public int LastRow { get; }
		public int LastColumn { get; }
	}

	internal sealed class WorksheetData
	{
		public WorksheetData(
			int index,
			string name,
			HeaderFooterText header,
			HeaderFooterText footer,
			int columnCount,
			IReadOnlyList<RowData> rows,
			IReadOnlyList<MergedRegion> mergedRegions)
		{
			Index = index;
			Name = name;
			Header = header ?? HeaderFooterText.Empty;
			Footer = footer ?? HeaderFooterText.Empty;
			ColumnCount = columnCount;
			Rows = rows;
			MergedRegions = mergedRegions;
		}

		public int Index { get; }
		public string Name { get; }
		public HeaderFooterText Header { get; }
		public HeaderFooterText Footer { get; }

		/// <summary>The width of the used area of the sheet in columns.</summary>
		public int ColumnCount { get; }

		public IReadOnlyList<RowData> Rows { get; }
		public IReadOnlyList<MergedRegion> MergedRegions { get; }

		public int LastRowIndex => Rows.Count == 0 ? 0 : Rows[Rows.Count - 1].RowIndex;
	}
}

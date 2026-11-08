using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Reads legacy binary workbooks (.xls) with ExcelDataReader.
	/// </summary>
	/// <remarks>
	/// ExcelDataReader is a forward-only reader which does not expose cell comments and formulas,
	/// so <see cref="CellData.Comment"/> and <see cref="CellData.Formula"/> are always <c>null</c> here.
	/// Sheets can only be read in ascending order.
	/// </remarks>
	internal sealed class ExcelDataReaderWorkbookReader : IWorkbookReader
	{
		private readonly Stream _stream;
		private readonly bool _ownsStream;
		private readonly IExcelDataReader _reader;
		private int _currentSheet;

		public ExcelDataReaderWorkbookReader(Stream stream, bool ownsStream)
		{
			_stream = stream;
			_ownsStream = ownsStream;
			_reader = ExcelReaderFactory.CreateBinaryReader(stream);
		}

		public int SheetCount => _reader.ResultsCount;

		public WorksheetData ReadSheet(int index)
		{
			if (index < _currentSheet)
			{
				throw new InvalidOperationException("The sheets of an .xls workbook can only be read in ascending order.");
			}
			while (_currentSheet < index)
			{
				if (!_reader.NextResult())
				{
					throw new ArgumentOutOfRangeException(nameof(index));
				}
				_currentSheet++;
			}

			var rows = new List<RowData>();
			int columnCount = 0;
			int rowIndex = -1;
			while (_reader.Read())
			{
				rowIndex++;
				columnCount = Math.Max(columnCount, _reader.FieldCount);
				var cells = new List<CellData>();
				for (int column = 0; column < _reader.FieldCount; column++)
				{
					object value = _reader.GetValue(column);
					if (value == null)
					{
						continue;
					}
					string text = CellValueFormatter.Format(value, _reader.GetNumberFormatString(column));
					cells.Add(new CellData(column, text, null, null, null));
				}
				rows.Add(new RowData(rowIndex, cells));
			}
			TrimEmptyRows(rows);

			var mergedRegions = new List<MergedRegion>();
			if (_reader.MergeCells != null)
			{
				foreach (CellRange range in _reader.MergeCells)
				{
					mergedRegions.Add(new MergedRegion(range.FromRow, range.FromColumn, range.ToRow, range.ToColumn));
				}
			}

			HeaderFooter headerFooter = _reader.HeaderFooter;
			return new WorksheetData(
				index,
				_reader.Name,
				HeaderFooterTextParser.Parse(headerFooter?.OddHeader),
				HeaderFooterTextParser.Parse(headerFooter?.OddFooter),
				columnCount,
				rows,
				mergedRegions);
		}

		public void Dispose()
		{
			_reader.Dispose();
			if (_ownsStream)
			{
				_stream.Dispose();
			}
		}

		/// <summary>
		/// The reader reports every row of the used area, which is also true for empty rows before and after the data.
		/// Only empty rows between data rows are kept, so the result matches the rows an .xlsx workbook reports.
		/// </summary>
		private static void TrimEmptyRows(List<RowData> rows)
		{
			while (rows.Count > 0 && rows[rows.Count - 1].Cells.Count == 0)
			{
				rows.RemoveAt(rows.Count - 1);
			}
			int leading = 0;
			while (leading < rows.Count && rows[leading].Cells.Count == 0)
			{
				leading++;
			}
			rows.RemoveRange(0, leading);
		}
	}
}

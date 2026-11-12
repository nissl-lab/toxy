using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using Toxy.Parsers.Excel;

namespace Toxy.Parsers
{
	/// <summary>
	/// Parses .xlsx and .xls workbooks with ExcelDataReader into a <see cref="ToxySpreadsheet"/>.
	/// </summary>
	/// <remarks>
	/// ExcelDataReader does not expose cell comments and formulas, so <c>IncludeComments</c> and
	/// <c>ShowCalculatedResult</c> have no effect.
	/// </remarks>
	public class ExcelSpreadsheetParser : ISpreadsheetParser
	{
		static ExcelSpreadsheetParser()
		{
			// .xls files store text in code pages
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
		}

		public ParserContext Context { get; set; }

		public ExcelSpreadsheetParser(ParserContext context)
		{
			Context = context;
		}

		public ToxySpreadsheet Parse()
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);

			ToxySpreadsheet ss = new ToxySpreadsheet();
			ss.Name = Context.IsStreamContext ? "Workbook1" : Context.Path;

			using (IExcelDataReader reader = CreateReader(Context))
			{
				for (int i = 0; i < reader.ResultsCount; i++)
				{
					if (i > 0)
					{
						reader.NextResult();
					}
					ss.Tables.Add(Parse(reader, i));
				}
			}
			return ss;
		}

		public ToxyTable Parse(int sheetIndex)
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);

			using (IExcelDataReader reader = CreateReader(Context))
			{
				if (sheetIndex < 0 || sheetIndex >= reader.ResultsCount)
				{
					throw new ArgumentOutOfRangeException(nameof(sheetIndex),
						string.Format("This file only contains {0} sheet(s).", reader.ResultsCount));
				}
				for (int i = 0; i < sheetIndex; i++)
				{
					reader.NextResult();
				}
				return Parse(reader, sheetIndex);
			}
		}

		/// <summary>
		/// Opens a reader for the workbook of <paramref name="context"/>. A stream passed in by the user stays open.
		/// </summary>
		private static IExcelDataReader CreateReader(ParserContext context)
		{
			if (context.IsStreamContext)
			{
				return CreateReader(context.Stream);
			}
			// the reader disposes the file when it is disposed itself
			FileStream file = File.OpenRead(context.Path);
			try
			{
				return ExcelReaderFactory.CreateReader(file);
			}
			catch
			{
				file.Dispose();
				throw;
			}
		}

		/// <summary>
		/// Opens a reader for an .xls or .xlsx workbook, <paramref name="stream"/> stays open.
		/// </summary>
		internal static IExcelDataReader CreateReader(Stream stream)
		{
			if (!stream.CanSeek)
			{
				// ExcelDataReader needs a seekable stream to detect the format
				MemoryStream buffer = new MemoryStream();
				stream.CopyTo(buffer);
				buffer.Position = 0;
				return ExcelReaderFactory.CreateReader(buffer);
			}
			stream.Position = 0;
			return ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration { LeaveOpen = true });
		}

		/// <summary>
		/// Reads the current sheet of <paramref name="reader"/>.
		/// </summary>
		private ToxyTable Parse(IExcelDataReader reader, int sheetIndex)
		{
			bool hasHeader = IsTrue("HasHeader", false);
			bool extractHeader = IsTrue("ExtractSheetHeader", false);
			bool extractFooter = IsTrue("ExtractSheetFooter", false);
			bool fillBlankCells = IsTrue("FillBlankCells", false);

			ToxyTable table = new ToxyTable();
			table.SheetIndex = sheetIndex;
			table.Name = reader.Name;

			// ExcelDataReader reports every row of the used area. Empty rows before the first and after the
			// last row with content are skipped, empty rows between them are kept.
			List<ToxyRow> emptyRows = new List<ToxyRow>();
			bool firstRow = true;
			int rowIndex = -1;
			while (reader.Read())
			{
				rowIndex++;
				ToxyRow row = new ToxyRow(rowIndex);
				bool hasValue = false;
				for (int column = 0; column < reader.FieldCount; column++)
				{
					object value = reader.GetValue(column);
					if (value == null && !fillBlankCells)
					{
						continue;
					}
					hasValue |= value != null;

					string text = CellValueFormatter.Format(reader, column);
					if (hasHeader && firstRow)
					{
						row.Cells.Add(new ToxyCell(column, text));
						continue;
					}
					if (row.LastCellIndex < column)
					{
						row.LastCellIndex = column;
					}
					if (text.Length > 0 || fillBlankCells)
					{
						row.Cells.Add(new ToxyCell(column, text));
					}
				}

				if (!hasValue)
				{
					if (!firstRow)
					{
						emptyRows.Add(row);
					}
					continue;
				}

				if (hasHeader && firstRow)
				{
					table.HeaderRows.Add(row);
				}
				else
				{
					AddRows(table, emptyRows);
					emptyRows.Clear();
					AddRows(table, new[] { row });
				}
				table.LastRowIndex = rowIndex;
				firstRow = false;
			}

			// header, footer and merged cells are known after all rows are read (.xlsx stores them behind the rows)
			if (extractHeader)
			{
				table.PageHeader = string.Join("|", HeaderFooterTextParser.Parse(reader.HeaderFooter?.OddHeader));
			}
			if (extractFooter)
			{
				table.PageFooter = string.Join("|", HeaderFooterTextParser.Parse(reader.HeaderFooter?.OddFooter));
			}
			if (reader.MergeCells != null)
			{
				foreach (CellRange range in reader.MergeCells)
				{
					table.MergeCells.Add(new MergeCellRange()
					{
						FirstRow = range.FromRow,
						FirstColumn = range.FromColumn,
						LastRow = range.ToRow,
						LastColumn = range.ToColumn
					});
				}
			}
			return table;
		}

		private static void AddRows(ToxyTable table, IEnumerable<ToxyRow> rows)
		{
			foreach (ToxyRow row in rows)
			{
				table.Rows.Add(row);
				if (table.LastColumnIndex < row.LastCellIndex)
				{
					table.LastColumnIndex = row.LastCellIndex;
				}
			}
		}

		private bool IsTrue(string key, bool defaultValue)
		{
			return Context.Properties.ContainsKey(key) ? Utility.IsTrue(Context.Properties[key]) : defaultValue;
		}
	}
}

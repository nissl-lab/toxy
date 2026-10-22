namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Maps a <see cref="WorksheetData"/> to a <see cref="ToxyTable"/>.
	/// </summary>
	internal static class ToxyTableBuilder
	{
		public static ToxyTable Build(WorksheetData sheet, ExcelParseOptions options)
		{
			ToxyTable table = new ToxyTable();
			table.SheetIndex = sheet.Index;
			table.Name = sheet.Name;
			table.LastRowIndex = sheet.LastRowIndex;

			if (options.ExtractSheetHeader)
			{
				table.PageHeader = Join(sheet.Header);
			}
			if (options.ExtractSheetFooter)
			{
				table.PageFooter = Join(sheet.Footer);
			}

			bool firstRow = true;
			foreach (RowData row in sheet.Rows)
			{
				bool isHeaderRow = options.HasHeader && firstRow;
				firstRow = false;

				ToxyRow target = new ToxyRow(row.RowIndex);
				for (int column = 0; column < sheet.ColumnCount; column++)
				{
					CellData cell = row.GetCell(column);
					if (cell == null && !options.FillBlankCells)
					{
						continue;
					}

					string text = cell != null ? cell.Text : string.Empty;
					if (isHeaderRow)
					{
						target.Cells.Add(new ToxyCell(column, text));
						continue;
					}

					if (target.LastCellIndex < column)
					{
						target.LastCellIndex = column;
					}
					ToxyCell toxyCell = new ToxyCell(column, text);
					if (text.Length > 0 || options.FillBlankCells)
					{
						target.Cells.Add(toxyCell);
					}
					if (options.IncludeComments && cell != null && cell.Comment != null)
					{
						toxyCell.Comment = cell.Comment;
					}
				}

				if (isHeaderRow)
				{
					table.HeaderRows.Add(target);
				}
				else
				{
					table.Rows.Add(target);
					if (table.LastColumnIndex < target.LastCellIndex)
					{
						table.LastColumnIndex = target.LastCellIndex;
					}
				}
			}

			foreach (MergedRegion region in sheet.MergedRegions)
			{
				table.MergeCells.Add(new MergeCellRange()
				{
					FirstRow = region.FirstRow,
					FirstColumn = region.FirstColumn,
					LastRow = region.LastRow,
					LastColumn = region.LastColumn
				});
			}
			return table;
		}

		private static string Join(HeaderFooterText text)
		{
			return text.Left + "|" + text.Center + "|" + text.Right;
		}
	}
}

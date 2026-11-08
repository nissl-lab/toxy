using System;
using Toxy.Parsers.Excel;

namespace Toxy.Parsers
{
	/// <summary>
	/// Parses .xlsx (Open XML SDK) and .xls (ExcelDataReader) workbooks into a <see cref="ToxySpreadsheet"/>.
	/// </summary>
	public class ExcelSpreadsheetParser : ISpreadsheetParser
	{
		public ParserContext Context { get; set; }

		public ExcelSpreadsheetParser(ParserContext context)
		{
			Context = context;
		}

		public ToxySpreadsheet Parse()
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);

			ExcelParseOptions options = ExcelParseOptions.From(Context);
			ToxySpreadsheet ss = new ToxySpreadsheet();
			ss.Name = Context.IsStreamContext ? "Workbook1" : Context.Path;

			using (IWorkbookReader workbook = WorkbookReaderFactory.Open(Context))
			{
				for (int i = 0; i < workbook.SheetCount; i++)
				{
					ss.Tables.Add(ToxyTableBuilder.Build(workbook.ReadSheet(i), options));
				}
			}
			return ss;
		}

		public ToxyTable Parse(int sheetIndex)
		{
			Utility.ValidateContext(Context);
			Utility.ThrowIfProtected(Context);

			ExcelParseOptions options = ExcelParseOptions.From(Context);
			using (IWorkbookReader workbook = WorkbookReaderFactory.Open(Context))
			{
				if (sheetIndex < 0 || sheetIndex >= workbook.SheetCount)
				{
					throw new ArgumentOutOfRangeException(nameof(sheetIndex),
						string.Format("This file only contains {0} sheet(s).", workbook.SheetCount));
				}
				return ToxyTableBuilder.Build(workbook.ReadSheet(sheetIndex), options);
			}
		}
	}
}

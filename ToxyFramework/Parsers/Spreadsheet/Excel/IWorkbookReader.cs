using System;

namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// Reads a spreadsheet workbook into the library independent model of <see cref="WorksheetData"/>.
	/// </summary>
	internal interface IWorkbookReader : IDisposable
	{
		/// <summary>The number of worksheets of the workbook.</summary>
		int SheetCount { get; }

		/// <summary>Reads the worksheet at the zero based <paramref name="index"/>.</summary>
		WorksheetData ReadSheet(int index);
	}
}

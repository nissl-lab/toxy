using System.IO;
using System.Text;
using Toxy.Base;
using Toxy.Parsers.Excel;

namespace Toxy.Parsers
{
	/// <summary>
	/// Extracts the text of .xlsx (Open XML SDK) and .xls (ExcelDataReader) workbooks.
	/// </summary>
	public class ExcelTextParser : BaseTextParser
	{
		public ExcelTextParser(ParserContext context) : base(context)
		{ }

		internal override void ValidateContext()
		{
			base.ValidateContext();
			Utility.ThrowIfProtected(Context);
		}

		internal override string ParseText(Stream stream)
		{
			ExcelParseOptions options = ExcelParseOptions.From(Context);
			StringBuilder builder = new StringBuilder();

			// the stream is disposed by BaseTextParser
			using (IWorkbookReader workbook = WorkbookReaderFactory.Open(stream, ownsStream: false))
			{
				for (int i = 0; i < workbook.SheetCount; i++)
				{
					ExcelTextWriter.AppendSheet(builder, workbook.ReadSheet(i), options);
				}
			}
			return builder.ToString();
		}
	}
}

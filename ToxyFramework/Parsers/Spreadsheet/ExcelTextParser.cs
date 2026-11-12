using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Toxy.Base;
using Toxy.Parsers.Excel;

namespace Toxy.Parsers
{
	/// <summary>
	/// Extracts the text of .xlsx and .xls workbooks with ExcelDataReader: one line per row, cells separated by tabs.
	/// </summary>
	/// <remarks>
	/// ExcelDataReader does not expose cell comments and formulas, so <c>IncludeComments</c> and
	/// <c>ShowCalculatedResult</c> have no effect.
	/// </remarks>
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
			bool extractHeaderFooter = false;
			if (Context.Properties.ContainsKey("IncludeHeaderFooter"))
			{
				extractHeaderFooter = Utility.IsTrue(Context.Properties["IncludeHeaderFooter"]);
			}
			bool includeSheetNames = true;
			if (Context.Properties.ContainsKey("IncludeSheetNames"))
			{
				includeSheetNames = Utility.IsTrue(Context.Properties["IncludeSheetNames"]);
			}

			StringBuilder builder = new StringBuilder();
			// the stream is disposed by BaseTextParser
			using (IExcelDataReader reader = ExcelSpreadsheetParser.CreateReader(stream))
			{
				do
				{
					if (includeSheetNames)
					{
						builder.Append(reader.Name).Append('\n');
					}
					StringBuilder rows = new StringBuilder();
					while (reader.Read())
					{
						List<string> cells = new List<string>(reader.FieldCount);
						for (int column = 0; column < reader.FieldCount; column++)
						{
							if (reader.GetValue(column) != null)
							{
								cells.Add(CellValueFormatter.Format(reader, column));
							}
						}
						AppendLine(rows, cells);
					}

					// the header and footer are known after all rows are read (.xlsx stores them behind the rows)
					if (extractHeaderFooter)
					{
						AppendLine(builder, HeaderFooterTextParser.Parse(reader.HeaderFooter?.OddHeader));
					}
					builder.Append(rows);
					if (extractHeaderFooter)
					{
						AppendLine(builder, HeaderFooterTextParser.Parse(reader.HeaderFooter?.OddFooter));
					}
				}
				while (reader.NextResult());
			}
			return builder.ToString();
		}

		/// <summary>
		/// Appends the non-empty <paramref name="texts"/> separated by tabs as one line, nothing if all are empty.
		/// </summary>
		private static void AppendLine(StringBuilder builder, IEnumerable<string> texts)
		{
			List<string> parts = new List<string>();
			foreach (string text in texts)
			{
				if (!string.IsNullOrEmpty(text))
				{
					parts.Add(text);
				}
			}
			if (parts.Count > 0)
			{
				builder.Append(string.Join("\t", parts)).Append('\n');
			}
		}
	}
}

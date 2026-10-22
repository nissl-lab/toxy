namespace Toxy.Parsers.Excel
{
	/// <summary>
	/// The options of the Excel parsers, read from <see cref="ParserContext.Properties"/>.
	/// </summary>
	internal sealed class ExcelParseOptions
	{
		public bool HasHeader { get; private set; }
		public bool ExtractSheetHeader { get; private set; }
		public bool ExtractSheetFooter { get; private set; }
		public bool ShowCalculatedResult { get; private set; }
		public bool FillBlankCells { get; private set; }
		public bool IncludeComments { get; private set; } = true;
		public bool IncludeHeaderFooter { get; private set; }
		public bool IncludeSheetNames { get; private set; } = true;

		public static ExcelParseOptions From(ParserContext context)
		{
			return new ExcelParseOptions
			{
				HasHeader = Read(context, "HasHeader", false),
				ExtractSheetHeader = Read(context, "ExtractSheetHeader", false),
				ExtractSheetFooter = Read(context, "ExtractSheetFooter", false),
				ShowCalculatedResult = Read(context, "ShowCalculatedResult", false),
				FillBlankCells = Read(context, "FillBlankCells", false),
				IncludeComments = Read(context, "IncludeComments", true),
				IncludeHeaderFooter = Read(context, "IncludeHeaderFooter", false),
				IncludeSheetNames = Read(context, "IncludeSheetNames", true),
			};
		}

		private static bool Read(ParserContext context, string key, bool defaultValue)
		{
			return context.Properties.ContainsKey(key) ? Utility.IsTrue(context.Properties[key]) : defaultValue;
		}
	}
}

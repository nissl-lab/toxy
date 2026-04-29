namespace Toxy.Parsers
{
    public class Excel2003ColumnarSpreadsheetParser : IColumnarSpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public Excel2003ColumnarSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxyColumnarSpreadsheet Parse()
        {
            // Default to treating the first row as column headers for columnar format
            if (!Context.Properties.ContainsKey("HasHeader"))
                Context.Properties["HasHeader"] = "1";

            var innerParser = new ExcelSpreadsheetParser(Context);
            ToxySpreadsheet ss = innerParser.Parse();
            return ToxyColumnarSpreadsheet.FromSpreadsheet(ss);
        }
    }
}

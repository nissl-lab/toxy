namespace Toxy.Parsers
{
    public class CsvColumnarSpreadsheetParser : IColumnarSpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public CsvColumnarSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxyColumnarSpreadsheet Parse()
        {
            // Default to extracting the first row as column headers for columnar format
            if (!Context.Properties.ContainsKey("ExtractHeader"))
                Context.Properties["ExtractHeader"] = "1";

            var innerParser = new CSVSpreadsheetParser(Context);
            ToxySpreadsheet ss = innerParser.Parse();
            return ToxyColumnarSpreadsheet.FromSpreadsheet(ss);
        }
    }
}

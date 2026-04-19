namespace Toxy
{
    public interface IColumnarSpreadsheetParser
    {
        ToxyColumnarSpreadsheet Parse();
        ParserContext Context { get; set; }
    }
}

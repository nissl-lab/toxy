namespace Toxy
{
    public interface ISpreadsheetParser
    {
        ToxySpreadsheet Parse();
        ToxyTable Parse(int sheetIndex);
        ParserContext Context { get; set; }
    }
}

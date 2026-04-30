using MsgReader.Outlook;
using System;
using Toxy.Helpers;

namespace Toxy.Parsers
{
    /// <summary>
    /// The <see cref="MsgSpreadsheetParser"/> extracts HTML and Markdown tables from a MSG message
    /// into a <see cref="ToxySpreadsheet"/>.
    /// </summary>
    public class MsgSpreadsheetParser : ISpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public MsgSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxySpreadsheet Parse()
        {
            Utility.ValidateContext(Context);

            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = Context.IsStreamContext ? "MsgDocument" : Context.Path;

            using (var stream = Utility.GetStream(Context))
            using (var reader = new Storage.Message(stream))
            {
                int tableOffset = 0;

                // Extract HTML tables from BodyHtml
                if (!string.IsNullOrEmpty(reader.BodyHtml))
                {
                    var htmlSpreadsheet = HtmlTableHelper.ParseFromHtml(reader.BodyHtml, "HtmlBody");
                    foreach (var table in htmlSpreadsheet.Tables)
                    {
                        table.SheetIndex = tableOffset++;
                        spreadsheet.Tables.Add(table);
                    }
                }

                // Extract Markdown tables from BodyText only if no HTML tables were found
                if (spreadsheet.Tables.Count == 0 && !string.IsNullOrEmpty(reader.BodyText))
                {
                    var mdSpreadsheet = MarkdownTableHelper.ParseFromMarkdown(reader.BodyText, "TextBody");
                    foreach (var table in mdSpreadsheet.Tables)
                    {
                        table.SheetIndex = tableOffset++;
                        spreadsheet.Tables.Add(table);
                    }
                }
            }

            return spreadsheet;
        }

        public ToxyTable Parse(int sheetIndex)
        {
            ToxySpreadsheet spreadsheet = Parse();
            if (sheetIndex < 0 || sheetIndex >= spreadsheet.Tables.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex),
                    $"Sheet index {sheetIndex} is out of range. Available sheets: {spreadsheet.Tables.Count}");
            }
            return spreadsheet.Tables[sheetIndex];
        }
    }
}

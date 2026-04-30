using MimeKit;
using System;
using System.IO;
using Toxy.Helpers;

namespace Toxy.Parsers
{
    /// <summary>
    /// The <see cref="EMLSpreadsheetParser"/> extracts HTML and Markdown tables from an EML message
    /// into a <see cref="ToxySpreadsheet"/>.
    /// </summary>
    public class EMLSpreadsheetParser : ISpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public EMLSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxySpreadsheet Parse()
        {
            Utility.ValidateContext(Context);

            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = Context.IsStreamContext ? "EmlDocument" : Context.Path;

            Stream stream = Utility.GetStream(Context);
            try
            {
                using (MimeMessage message = MimeMessage.Load(stream))
                {
                    int tableOffset = 0;

                    // Extract HTML tables from HtmlBody
                    if (!string.IsNullOrEmpty(message.HtmlBody))
                    {
                        var htmlSpreadsheet = HtmlTableHelper.ParseFromHtml(message.HtmlBody, "HtmlBody");
                        foreach (var table in htmlSpreadsheet.Tables)
                        {
                            table.SheetIndex = tableOffset++;
                            spreadsheet.Tables.Add(table);
                        }
                    }

                    // Extract Markdown tables from TextBody only if no HTML tables were found
                    if (spreadsheet.Tables.Count == 0 && !string.IsNullOrEmpty(message.TextBody))
                    {
                        var mdSpreadsheet = MarkdownTableHelper.ParseFromMarkdown(message.TextBody, "TextBody");
                        foreach (var table in mdSpreadsheet.Tables)
                        {
                            table.SheetIndex = tableOffset++;
                            spreadsheet.Tables.Add(table);
                        }
                    }
                }
            }
            finally
            {
                if (!Context.IsStreamContext)
                    stream.Dispose();
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

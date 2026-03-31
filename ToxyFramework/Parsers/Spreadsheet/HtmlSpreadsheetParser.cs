using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Toxy.Parsers
{
    public class HtmlSpreadsheetParser : ISpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public HtmlSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxySpreadsheet Parse()
        {
            Utility.ValidateContext(Context);

            HtmlDocument htmlDoc;
            if (Context.IsStreamContext)
            {
                htmlDoc = new HtmlDocument();
                htmlDoc.Load(Context.Stream);
            }
            else
            {
                htmlDoc = new HtmlDocument();
                htmlDoc.Load(Context.Path);
            }

            if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Any())
            {
                var firstError = htmlDoc.ParseErrors.First();
                throw new HtmlParseException(
                    $"HTML parsing error at line {firstError.Line}, column {firstError.LinePosition}: {firstError.Reason}");
            }

            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = Context.IsStreamContext ? "HtmlDocument" : Context.Path;

            var tableNodes = htmlDoc.DocumentNode.SelectNodes("//table");
            if (tableNodes != null)
            {
                int tableIndex = 0;
                foreach (var tableNode in tableNodes)
                {
                    ToxyTable table = ParseTable(tableNode, tableIndex);
                    spreadsheet.Tables.Add(table);
                    tableIndex++;
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

        private ToxyTable ParseTable(HtmlNode tableNode, int tableIndex)
        {
            ToxyTable table = new ToxyTable();
            table.SheetIndex = tableIndex;
            table.Name = GetTableName(tableNode, tableIndex);

            var rows = tableNode.SelectNodes(".//tr");
            if (rows == null)
            {
                return table;
            }

            int rowIndex = 0;
            foreach (var rowNode in rows)
            {
                ToxyRow row = new ToxyRow(rowIndex);
                var cells = rowNode.SelectNodes("td|th");
                if (cells != null)
                {
                    int cellIndex = 0;
                    foreach (var cellNode in cells)
                    {
                        string cellValue = GetCellValue(cellNode);
                        ToxyCell cell = new ToxyCell(cellIndex, cellValue);
                        row.Cells.Add(cell);
                        cellIndex++;
                    }
                    row.LastCellIndex = cellIndex - 1;
                    if (row.LastCellIndex > table.LastColumnIndex)
                    {
                        table.LastColumnIndex = row.LastCellIndex;
                    }
                }
                table.Rows.Add(row);
                rowIndex++;
            }
            table.LastRowIndex = rowIndex - 1;

            return table;
        }

        private string GetTableName(HtmlNode tableNode, int tableIndex)
        {
            var idAttr = tableNode.GetAttributeValue("id", null);
            if (!string.IsNullOrEmpty(idAttr))
            {
                return idAttr;
            }

            var nameAttr = tableNode.GetAttributeValue("name", null);
            if (!string.IsNullOrEmpty(nameAttr))
            {
                return nameAttr;
            }

            return $"Table_{tableIndex + 1}";
        }

        private string GetCellValue(HtmlNode cellNode)
        {
            return cellNode.InnerText?.Trim() ?? string.Empty;
        }
    }

    public class HtmlParseException : Exception
    {
        public HtmlParseException(string message) : base(message) { }
    }
}

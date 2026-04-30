using HtmlAgilityPack;
using System.Collections.Generic;

namespace Toxy.Helpers
{
    internal static class HtmlTableHelper
    {
        internal static ToxySpreadsheet ParseFromHtml(string htmlContent, string name)
        {
            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = name;

            if (string.IsNullOrEmpty(htmlContent))
                return spreadsheet;

            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);

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

        private static ToxyTable ParseTable(HtmlNode tableNode, int tableIndex)
        {
            ToxyTable table = new ToxyTable();
            table.SheetIndex = tableIndex;
            table.Name = GetTableName(tableNode, tableIndex);

            var rows = tableNode.SelectNodes(".//tr");
            if (rows == null)
                return table;

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
                        string cellValue = cellNode.InnerText?.Trim() ?? string.Empty;
                        ToxyCell cell = new ToxyCell(cellIndex, cellValue);
                        row.Cells.Add(cell);
                        cellIndex++;
                    }
            row.LastCellIndex = cellIndex > 0 ? cellIndex - 1 : 0;
            if (row.LastCellIndex > table.LastColumnIndex)
                table.LastColumnIndex = row.LastCellIndex;
                }
                table.Rows.Add(row);
                rowIndex++;
            }
            table.LastRowIndex = rowIndex - 1;

            return table;
        }

        private static string GetTableName(HtmlNode tableNode, int tableIndex)
        {
            var idAttr = tableNode.GetAttributeValue("id", null);
            if (!string.IsNullOrEmpty(idAttr))
                return idAttr;

            var nameAttr = tableNode.GetAttributeValue("name", null);
            if (!string.IsNullOrEmpty(nameAttr))
                return nameAttr;

            return $"Table_{tableIndex + 1}";
        }
    }
}

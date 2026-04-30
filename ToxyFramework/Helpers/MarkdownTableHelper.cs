using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Toxy.Helpers
{
    internal static class MarkdownTableHelper
    {
        private static readonly Regex SeparatorRegex = new Regex(@"^\|?[\s\-:]+(\|[\s\-:]+)*\|?$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        internal static ToxySpreadsheet ParseFromMarkdown(string markdownText, string name)
        {
            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = name;

            if (string.IsNullOrEmpty(markdownText))
                return spreadsheet;

            string[] lines = markdownText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            int tableIndex = 0;

            int i = 0;
            while (i < lines.Length)
            {
                if (IsPipeRow(lines[i]))
                {
                    // Collect all contiguous pipe rows into a block
                    int blockStart = i;
                    while (i < lines.Length && IsPipeRow(lines[i]))
                    {
                        i++;
                    }
                    int blockEnd = i - 1; // inclusive

                    ToxyTable table = ParseTableBlock(lines, blockStart, blockEnd, tableIndex);
                    if (table.Rows.Count > 0 || table.HeaderRows.Count > 0)
                    {
                        spreadsheet.Tables.Add(table);
                        tableIndex++;
                    }
                }
                else
                {
                    i++;
                }
            }

            return spreadsheet;
        }

        private static bool IsPipeRow(string line)
        {
            return line.IndexOf('|') >= 0;
        }

        private static ToxyTable ParseTableBlock(string[] lines, int blockStart, int blockEnd, int tableIndex)
        {
            ToxyTable table = new ToxyTable();
            table.SheetIndex = tableIndex;
            table.Name = $"Table_{tableIndex + 1}";

            // Find separator row position within the block
            int separatorIndex = -1;
            for (int i = blockStart; i <= blockEnd; i++)
            {
                if (SeparatorRegex.IsMatch(lines[i].Trim()))
                {
                    separatorIndex = i;
                    break;
                }
            }

            int rowIndex = 0;

            if (separatorIndex == blockStart + 1)
            {
                // Standard: header row is blockStart, separator is blockStart+1, data starts at blockStart+2
                ToxyRow headerRow = ParsePipeRow(lines[blockStart], 0);
                if (headerRow != null)
                    table.HeaderRows.Add(headerRow);

                for (int i = blockStart + 2; i <= blockEnd; i++)
                {
                    ToxyRow row = ParsePipeRow(lines[i], rowIndex);
                    if (row != null)
                    {
                        UpdateTableMetrics(table, row);
                        table.Rows.Add(row);
                        rowIndex++;
                    }
                }
            }
            else
            {
                // No separator row (or separator not in expected position) — treat all rows as data
                for (int i = blockStart; i <= blockEnd; i++)
                {
                    if (SeparatorRegex.IsMatch(lines[i].Trim()))
                        continue; // skip separator rows

                    ToxyRow row = ParsePipeRow(lines[i], rowIndex);
                    if (row != null)
                    {
                        UpdateTableMetrics(table, row);
                        table.Rows.Add(row);
                        rowIndex++;
                    }
                }
            }

            if (rowIndex > 0)
                table.LastRowIndex = rowIndex - 1;

            return table;
        }

        private static void UpdateTableMetrics(ToxyTable table, ToxyRow row)
        {
            if (row.LastCellIndex > table.LastColumnIndex)
                table.LastColumnIndex = row.LastCellIndex;
        }

        private static ToxyRow ParsePipeRow(string line, int rowIndex)
        {
            string trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed))
                return null;

            // Remove leading and trailing pipe if present
            if (trimmed.StartsWith("|"))
                trimmed = trimmed.Substring(1);
            if (trimmed.EndsWith("|"))
                trimmed = trimmed.Substring(0, trimmed.Length - 1);

            string[] parts = trimmed.Split('|');
            ToxyRow row = new ToxyRow(rowIndex);
            int cellIndex = 0;
            foreach (string part in parts)
            {
                string cellValue = part.Trim();
                ToxyCell cell = new ToxyCell(cellIndex, cellValue);
                row.Cells.Add(cell);
                cellIndex++;
            }
            row.LastCellIndex = cellIndex > 0 ? cellIndex - 1 : 0;
            return row;
        }
    }
}

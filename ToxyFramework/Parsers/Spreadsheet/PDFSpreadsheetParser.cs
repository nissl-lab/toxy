using System;
using System.Collections.Generic;
using System.IO;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig;

namespace Toxy.Parsers
{
    public class PDFSpreadsheetParser : ISpreadsheetParser
    {
        public ParserContext Context { get; set; }

        public PDFSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public ToxySpreadsheet Parse()
        {
            Utility.ValidateContext(Context);
            Utility.ThrowIfProtected(Context);

            ToxySpreadsheet spreadsheet = new ToxySpreadsheet();
            spreadsheet.Name = Context.IsStreamContext ? "PDFTables" : Path.GetFileName(Context.Path);

            Stream stream = Utility.GetStream(Context);
            try
            {
                using (PdfDocument document = PdfDocument.Open(stream))
                {
                    var extractionAlgorithm = new SpreadsheetExtractionAlgorithm();

                    int tableIndex = 0;
                    int pageNumber = 1;
                    foreach (var page in document.GetPages())
                    {
                        var pageArea = ObjectExtractor.Extract(document, pageNumber);
                        var tables = extractionAlgorithm.Extract(pageArea);

                        foreach (var table in tables)
                        {
                            ToxyTable toxyTable = ConvertTable(table, tableIndex, pageNumber);
                            spreadsheet.Tables.Add(toxyTable);
                            tableIndex++;
                        }
                        pageNumber++;
                    }
                }
            }
            finally
            {
                if (!Context.IsStreamContext)
                {
                    stream.Dispose();
                }
            }

            return spreadsheet;
        }

        public ToxyTable Parse(int sheetIndex)
        {
            var spreadsheet = Parse();
            if (sheetIndex < 0 || sheetIndex >= spreadsheet.Tables.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), 
                    $"Sheet index {sheetIndex} is out of range. Available tables: {spreadsheet.Tables.Count}");
            }
            return spreadsheet.Tables[sheetIndex];
        }

        private ToxyTable ConvertTable(Table table, int tableIndex, int pageNumber)
        {
            ToxyTable toxyTable = new ToxyTable();
            toxyTable.Name = $"Table {tableIndex + 1} (Page {pageNumber})";
            toxyTable.SheetIndex = tableIndex;

            int rowIndex = 0;
            foreach (var row in table.Rows)
            {
                ToxyRow toxyRow = new ToxyRow(rowIndex);
                toxyRow.RowIndex = rowIndex;

                int cellIndex = 0;
                foreach (var cell in row)
                {
                    string cellValue = cell?.ToString() ?? string.Empty;
                    ToxyCell toxyCell = new ToxyCell(cellIndex, cellValue);
                    toxyRow.Cells.Add(toxyCell);
                    cellIndex++;
                }

                if (toxyRow.Cells.Count > 0)
                {
                    toxyRow.LastCellIndex = toxyRow.Cells.Count - 1;
                    if (toxyTable.LastColumnIndex < toxyRow.LastCellIndex)
                    {
                        toxyTable.LastColumnIndex = toxyRow.LastCellIndex;
                    }
                }

                toxyTable.Rows.Add(toxyRow);
                rowIndex++;
            }

            toxyTable.LastRowIndex = rowIndex - 1;
            return toxyTable;
        }
    }
}

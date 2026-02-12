using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;

namespace Toxy.Parsers
{
    public class ExcelSpreadsheetParser : ISpreadsheetParser
    {
        public ExcelSpreadsheetParser(ParserContext context)
        {
            this.Context = context;
        }
        public ToxySpreadsheet Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            bool hasHeader = false;
            if (Context.Properties.ContainsKey("HasHeader"))
            {
                hasHeader = Utility.IsTrue(Context.Properties["HasHeader"]);
            }
            bool extractHeader = false;
            if (Context.Properties.ContainsKey("ExtractSheetHeader"))
            {
                extractHeader = Utility.IsTrue(Context.Properties["ExtractSheetHeader"]);
            }
            bool extractFooter = false;
            if (Context.Properties.ContainsKey("ExtractSheetFooter"))
            {
                extractFooter = Utility.IsTrue(Context.Properties["ExtractSheetFooter"]);
            }
            bool showCalculatedResult = false;
            if (Context.Properties.ContainsKey("ShowCalculatedResult"))
            {
                showCalculatedResult = Utility.IsTrue(Context.Properties["ShowCalculatedResult"]);
            }
            bool fillBlankCells = false;
            if (Context.Properties.ContainsKey("FillBlankCells"))
            {
                fillBlankCells = Utility.IsTrue(Context.Properties["FillBlankCells"]);
            }
            bool includeComment = true;
            if (Context.Properties.ContainsKey("IncludeComments"))
            {
                includeComment = Utility.IsTrue(Context.Properties["IncludeComments"]);
            }
            ToxySpreadsheet ss = new ToxySpreadsheet();
            ss.Name = Context.Path;
            IWorkbook workbook = WorkbookFactory.Create(Context.Path);
           
            HSSFDataFormatter formatter = new HSSFDataFormatter();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ToxyTable table = Parse(workbook, i, extractHeader, extractFooter, hasHeader, fillBlankCells, includeComment, formatter);
                ss.Tables.Add(table);
            }
            return ss;
        }

        public ParserContext Context
        {
            get;
            set;
        }
        ToxyTable Parse(IWorkbook workbook, int sheetIndex, bool extractHeader, bool extractFooter, bool hasHeader, bool fillBlankCells, bool includeComment, HSSFDataFormatter formatter)
        {
            ToxyTable table = new ToxyTable();
            if (workbook.NumberOfSheets - 1 < sheetIndex)
            {
                throw new ArgumentOutOfRangeException(string.Format("This file only contains {0} sheet(s).", workbook.NumberOfSheets));
            }
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            table.SheetIndex = sheetIndex;
            table.Name = sheet.SheetName;

            if (extractHeader && sheet.Header != null)
            {
                table.PageHeader = sheet.Header.Left + "|" + sheet.Header.Center + "|" + sheet.Header.Right;
            }

            if (extractFooter && sheet.Footer != null)
            {
                table.PageFooter = sheet.Footer.Left + "|" + sheet.Footer.Center + "|" + sheet.Footer.Right;
            }

            bool firstRow = true;
            table.LastRowIndex = sheet.LastRowNum;
            foreach (IRow row in sheet)
            {
                ToxyRow tr = null;
                if (!hasHeader || !firstRow)
                {
                    tr = new ToxyRow(row.RowNum);
                }
                else if (hasHeader && firstRow)
                {
                    table.HeaderRows.Add(new ToxyRow(row.RowNum));
                }
                for (int i=0;i<row.LastCellNum;i++)
                {
                    ICell cell = row.GetCell(i, fillBlankCells?MissingCellPolicy.CREATE_NULL_AS_BLANK:MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    if (hasHeader && firstRow)
                    {
                        table.HeaderRows[0].Cells.Add(new ToxyCell(cell.ColumnIndex, cell.ToString()));
                    }
                    else if(cell!=null)
                    {
                        if (tr.LastCellIndex < cell.ColumnIndex)
                        {
                            tr.LastCellIndex = cell.ColumnIndex;
                        }
                        ToxyCell c = new ToxyCell(cell.ColumnIndex, formatter.FormatCellValue(cell));
                        if (!string.IsNullOrEmpty(cell.ToString()))
                        {
                            tr.Cells.Add(c);
                        }
                        else if (fillBlankCells)
                        {
                            tr.Cells.Add(c);
                        }
                        if (includeComment && cell.CellComment != null)
                        {
                            c.Comment = cell.CellComment.String.String;
                        }
                    }
                }
                if (tr != null)
                {
                    tr.RowIndex = row.RowNum;
                    table.Rows.Add(tr);

                    if (table.LastColumnIndex < tr.LastCellIndex)
                        table.LastColumnIndex = tr.LastCellIndex;
                }
                if (firstRow)
                {
                    firstRow = false;
                }
            }
            for (int j = 0; j < sheet.NumMergedRegions; j++)
            {
                var range = sheet.GetMergedRegion(j);
                table.MergeCells.Add(new MergeCellRange() { FirstRow = range.FirstRow, FirstColumn = range.FirstColumn, LastRow = range.LastRow, LastColumn = range.LastColumn });
            }
            return table;
        }

        public ToxyTable Parse(int sheetIndex)
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            bool hasHeader = false;
            if (Context.Properties.ContainsKey("HasHeader"))
            {
                hasHeader = Utility.IsTrue(Context.Properties["HasHeader"]);
            }
            bool extractHeader = false;
            if (Context.Properties.ContainsKey("ExtractSheetHeader"))
            {
                extractHeader = Utility.IsTrue(Context.Properties["ExtractSheetHeader"]);
            }
            bool extractFooter = false;
            if (Context.Properties.ContainsKey("ExtractSheetFooter"))
            {
                extractFooter = Utility.IsTrue(Context.Properties["ExtractSheetFooter"]);
            }
            bool showCalculatedResult = false;
            if (Context.Properties.ContainsKey("ShowCalculatedResult"))
            {
                showCalculatedResult = Utility.IsTrue(Context.Properties["ShowCalculatedResult"]);
            }
            bool fillBlankCells = false;
            if (Context.Properties.ContainsKey("FillBlankCells"))
            {
                fillBlankCells = Utility.IsTrue(Context.Properties["FillBlankCells"]);
            }
            bool includeComment = true;
            if (Context.Properties.ContainsKey("IncludeComments"))
            {
                includeComment = Utility.IsTrue(Context.Properties["IncludeComments"]);
            }
            IWorkbook workbook = WorkbookFactory.Create(Context.Path);

            HSSFDataFormatter formatter = new HSSFDataFormatter();
            return Parse(workbook, sheetIndex, extractHeader, extractFooter, hasHeader, fillBlankCells, includeComment, formatter);
        }
    }
}

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

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
            if (Context.Properties.ContainsKey("GenerateColumnHeader"))
            {
                hasHeader = Utility.IsTrue(Context.Properties["GenerateColumnHeader"]);
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
                ToxyTable table=new ToxyTable();
                ISheet sheet = workbook.GetSheetAt(i);
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
                    ToxyRow tr=null;
                    if (!hasHeader || !firstRow)
                    {
                        tr=new ToxyRow(row.RowNum);
                    }
                    foreach (ICell cell in row)
                    {
                        if (hasHeader&& firstRow)
                        {
                            table.ColumnHeaders.Cells.Add(new ToxyCell(cell.ColumnIndex, cell.ToString()));
                        }
                        else 
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
                            if (cell.CellComment != null)
                            {
                                c.Comment = cell.CellComment.String.String;
                            }
                        }
                    }
                    if (tr != null)
                    {
                        tr.RowIndex = row.RowNum;
                        table.Rows.Add(tr);
                    }
                    if (firstRow)
                    {
                        firstRow = false;
                    }
                    if(table.LastColumnIndex<tr.LastCellIndex)
                        table.LastColumnIndex=tr.LastCellIndex;
                }
                for (int j = 0; j < sheet.NumMergedRegions; j++)
                { 
                    var range = sheet.GetMergedRegion(j);
                    table.MergeCells.Add(new MergeCellRange() { FirstRow = range.FirstRow, FirstColumn = range.FirstColumn, LastRow = range.LastRow, LastColumn = range.LastColumn });
                }
                ss.Tables.Add(table);
            }
            if (workbook is XSSFWorkbook)
            {
                var props= ((XSSFWorkbook)workbook).GetProperties();

                if (props.CoreProperties != null)
                {
                    if (props.CoreProperties.Title != null)
                    {
                        ss.Properties.Add("Title", props.CoreProperties.Title );
                    }
                    else if (props.CoreProperties.Identifier != null)
                    {
                        ss.Properties.Add("Identifier", props.CoreProperties.Identifier );
                    }
                    else if (props.CoreProperties.Keywords != null)
                    {
                        ss.Properties.Add("Keywords", props.CoreProperties.Keywords);
                    }
                    else if (props.CoreProperties.Revision != null)
                    {
                        ss.Properties.Add("Revision", props.CoreProperties.Revision);
                    }
                    else if (props.CoreProperties.Subject != null)
                    {
                        ss.Properties.Add("Subject", props.CoreProperties.Subject);
                    }
                    else if (props.CoreProperties.Modified != null)
                    {
                        ss.Properties.Add("Modified", props.CoreProperties.Modified);
                    }
                    else if (props.CoreProperties.LastPrinted != null)
                    {
                        ss.Properties.Add("LastPrinted", props.CoreProperties.LastPrinted);
                    }
                    else if (props.CoreProperties.Created != null)
                    {
                        ss.Properties.Add("Created", props.CoreProperties.Created);
                    }
                    else if (props.CoreProperties.Creator != null)
                    {
                        ss.Properties.Add("Creator", props.CoreProperties.Creator);
                    }
                    else if (props.CoreProperties.Description != null)
                    {
                        ss.Properties.Add("Description", props.CoreProperties.Description);
                    }
                }
                if (props.ExtendedProperties != null && props.ExtendedProperties.props!=null)
                {
                    var extProps = props.ExtendedProperties.props.GetProperties();
                    if (extProps.Application != null)
                    {
                        ss.Properties.Add("Application", extProps.Application);
                    }
                    if (extProps.AppVersion != null)
                    {
                        ss.Properties.Add("AppVersion", extProps.AppVersion);
                    }
                    if (extProps.Characters>0)
                    {
                        ss.Properties.Add("Characters", extProps.Characters);
                    }
                    if (extProps.CharactersWithSpaces>0)
                    {
                        ss.Properties.Add("CharactersWithSpaces", extProps.CharactersWithSpaces);
                    }
                    if (extProps.Company != null)
                    {
                        ss.Properties.Add("Company", extProps.Company);
                    }
                    if (extProps.Lines > 0)
                    {
                        ss.Properties.Add("Lines", extProps.Lines);
                    }
                    if (extProps.Manager != null)
                    {
                        ss.Properties.Add("Manager", extProps.Manager);
                    }
                    if (extProps.Notes> 0)
                    {
                        ss.Properties.Add("Notes", extProps.Notes);
                    }
                    if (extProps.Pages>0)
                    {
                        ss.Properties.Add("Pages", extProps.Pages);
                    }
                    if (extProps.Paragraphs>0)
                    {
                        ss.Properties.Add("Paragraphs", extProps.Paragraphs);
                    }
                    if (extProps.Words>0)
                    {
                        ss.Properties.Add("Words", extProps.Words);
                    }
                    if (extProps.TotalTime>0)
                    {
                        ss.Properties.Add("TotalTime", extProps.TotalTime);
                    }
                }
            }
            else
            {
                //HSSFWorkbook
                var si = ((HSSFWorkbook)workbook).SummaryInformation;
                if (si != null)
                {
                    if (si.Title != null)
                    {
                        ss.Properties.Add("Title", si.Title);
                    }
                    else if (si.LastSaveDateTime != null)
                    {
                        ss.Properties.Add("LastSaveDateTime", si.LastSaveDateTime);
                    }
                    else if (si.PageCount > 0)
                    {
                        ss.Properties.Add("PageCount", si.PageCount);
                    }
                    else if (si.OSVersion > 0)
                    {
                        ss.Properties.Add("OSVersion", si.OSVersion);
                    }
                    else if (si.Security > 0)
                    {
                        ss.Properties.Add("Security", si.Security);
                    }
                    else if (si.Keywords != null)
                    {
                        ss.Properties.Add("Keywords", si.Keywords);
                    }
                    else if (si.EditTime > 0)
                    {
                        ss.Properties.Add("EditTime", si.EditTime);
                    }
                    else if (si.Subject != null)
                    {
                        ss.Properties.Add("Subject", si.Subject);
                    }
                    else if (si.CreateDateTime != null)
                    {
                        ss.Properties.Add("CreateDateTime", si.CreateDateTime);
                    }
                    else if (si.LastPrinted != null)
                    {
                        ss.Properties.Add("LastPrinted", si.LastPrinted);
                    }
                    else if (si.CharCount != null)
                    {
                        ss.Properties.Add("CharCount", si.CharCount);
                    }
                    else if (si.Author != null)
                    {
                        ss.Properties.Add("Author", si.Author);
                    }
                    else if (si.LastAuthor != null)
                    {
                        ss.Properties.Add("LastAuthor", si.LastAuthor);
                    }
                    else if (si.ApplicationName != null)
                    {
                        ss.Properties.Add("ApplicationName", si.ApplicationName);
                    }
                    else if (si.RevNumber != null)
                    {
                        ss.Properties.Add("RevNumber", si.RevNumber);
                    }
                    else if (si.Template != null)
                    {
                        ss.Properties.Add("Template", si.Template);
                    }
                }
                var dsi = ((HSSFWorkbook)workbook).DocumentSummaryInformation;
                if(dsi!=null)
                {
                    if (dsi.ByteCount > 0)
                    {
                        ss.Properties.Add("ByteCount", dsi.ByteCount);
                    }
                    else if (dsi.Company !=null)
                    {
                        ss.Properties.Add("Company", dsi.Company);
                    }
                    else if (dsi.Format>0)
                    {
                        ss.Properties.Add("Format", dsi.Format);
                    }
                    else if (dsi.LineCount!= null)
                    {
                        ss.Properties.Add("LineCount", dsi.Company);
                    }
                    else if (dsi.LinksDirty)
                    {
                        ss.Properties.Add("LinksDirty", true);
                    }
                    else if (dsi.Manager!=null)
                    {
                        ss.Properties.Add("Manager", dsi.Manager);
                    }
                    else if (dsi.NoteCount != null)
                    {
                        ss.Properties.Add("NoteCount", dsi.NoteCount);
                    }
                    else if (dsi.Scale)
                    {
                        ss.Properties.Add("Scale", dsi.Scale);
                    }
                    else if (dsi.Company != null)
                    {
                        ss.Properties.Add("Company", dsi.Company);
                    }
                    else if (dsi.MMClipCount != null)
                    {
                        ss.Properties.Add("MMClipCount", dsi.MMClipCount);
                    }
                    else if (dsi.ParCount != null)
                    {
                        ss.Properties.Add("ParCount", dsi.ParCount);
                    }
                }
            }
            return ss;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

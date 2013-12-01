using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class ExcelParser : ISpreadsheetParser
    {
        public ToxySpreadsheet Parse(ParserContext context)
        {
            if (!File.Exists(context.Path))
                throw new FileNotFoundException("File " + context.Path + " is not found");

            bool hasHeader = false;
            if (context.Properties.ContainsKey("HasColumnHeader"))
            {
                hasHeader = Utility.IsTrue(context.Properties["HasColumnHeader"]);
            }
            bool extractHeader = false;
            if (context.Properties.ContainsKey("ExtractSheetHeader"))
            {
                hasHeader = Utility.IsTrue(context.Properties["ExtractSheetHeader"]);
            }
            bool extractFooter = false;
            if (context.Properties.ContainsKey("ExtractFooter"))
            {
                extractFooter = Utility.IsTrue(context.Properties["ExtractFooter"]);
            }
            bool showCalculatedResult = false;
            if (context.Properties.ContainsKey("ShowCalculatedResult"))
            {
                showCalculatedResult = Utility.IsTrue(context.Properties["ShowCalculatedResult"]);
            }
            ToxySpreadsheet ss = new ToxySpreadsheet();
            IWorkbook workbook = WorkbookFactory.Create(context.Path);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ToxyTable table=new ToxyTable();
                ISheet sheet = workbook.GetSheetAt(i);
                table.Name = sheet.SheetName;
                if (extractHeader && sheet.Header != null)
                {
                    table.PageHeader = sheet.Header.Left + " | " + sheet.Header.Center + " | " + sheet.Header.Right;
                }

                if (extractFooter && sheet.Footer != null)
                {
                    table.PageFooter = sheet.Footer.Left + " | " + sheet.Footer.Center + " | " + sheet.Footer.Right;
                }
                foreach (IRow row in sheet)
                {
                    ToxyRow tr=null;
                    if(!hasHeader)
                    {
                        tr=new ToxyRow();
                    }

                    foreach (ICell cell in row)
                    {
                        if (hasHeader&& row.RowNum == 0)
                        {
                            table.ColumnHeaders.Add(cell.ToString());
                        }
                        else
                        {
                            tr.Cells.Add(cell.ToString());
                        }
                    }
                    if (tr != null)
                    {
                        table.Rows.Add(tr);
                    }
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
                    if (extProps.Characters != null)
                    {
                        ss.Properties.Add("Characters", extProps.Characters);
                    }
                    if (extProps.CharactersWithSpaces != null)
                    {
                        ss.Properties.Add("CharactersWithSpaces", extProps.CharactersWithSpaces);
                    }
                    if (extProps.Company != null)
                    {
                        ss.Properties.Add("Company", extProps.Company);
                    }
                    if (extProps.Lines != null)
                    {
                        ss.Properties.Add("Lines", extProps.Lines);
                    }
                    if (extProps.Manager != null)
                    {
                        ss.Properties.Add("Manager", extProps.Manager);
                    }
                    if (extProps.Notes != null)
                    {
                        ss.Properties.Add("Notes", extProps.Notes);
                    }
                    if (extProps.Pages != null)
                    {
                        ss.Properties.Add("Pages", extProps.Pages);
                    }
                    if (extProps.Paragraphs != null)
                    {
                        ss.Properties.Add("Paragraphs", extProps.Paragraphs);
                    }
                    if (extProps.Words != null)
                    {
                        ss.Properties.Add("Words", extProps.Words);
                    }
                    if (extProps.TotalTime != null)
                    {
                        ss.Properties.Add("TotalTime", extProps.TotalTime);
                    }
                }
            }
            else
            {
                //HSSFWorkbook
                throw new NotImplementedException();
            }
            return ss;
        }
    }
}

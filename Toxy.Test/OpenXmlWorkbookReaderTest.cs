using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;

namespace Toxy.Test
{
    /// <summary>
    /// Tests for .xlsx features which are not covered by the files in testdata, the workbooks are created in memory.
    /// </summary>
    [TestFixture]
    public class OpenXmlWorkbookReaderTest
    {
        [Test]
        public void TestSharedFormula()
        {
            // B1:B2 share the formula of B1, only B1 stores the formula text
            MemoryStream stream = CreateWorkbook(false, null, sheetData =>
            {
                sheetData.Append(new Row(
                    new Cell { CellReference = "A1", CellValue = new CellValue("1") },
                    new Cell
                    {
                        CellReference = "B1",
                        CellFormula = new CellFormula("A1*2") { FormulaType = CellFormulaValues.Shared, Reference = "B1:B2", SharedIndex = 0 },
                        CellValue = new CellValue("2")
                    }) { RowIndex = 1 });
                sheetData.Append(new Row(
                    new Cell { CellReference = "A2", CellValue = new CellValue("2") },
                    new Cell
                    {
                        CellReference = "B2",
                        CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 },
                        CellValue = new CellValue("4")
                    }) { RowIndex = 2 });
            });

            ParserContext context = new ParserContext(stream);
            context.Properties.Add("IncludeSheetNames", "0");
            string result = ParserFactory.CreateText(context).Parse();
            ClassicAssert.IsTrue(result.Contains("1\tA1*2"));
            ClassicAssert.IsTrue(result.Contains("2\t4"));

            context.Properties.Add("ShowCalculatedResult", "1");
            result = ParserFactory.CreateText(context).Parse();
            ClassicAssert.IsTrue(result.Contains("1\t2"));
            ClassicAssert.IsTrue(result.Contains("2\t4"));
        }

        [Test]
        public void TestDate1904()
        {
            // style 1 uses the built-in format 14 (mm-dd-yy)
            CellFormats cellFormats = new CellFormats(
                new CellFormat { NumberFormatId = 0 },
                new CellFormat { NumberFormatId = 14, ApplyNumberFormat = true });
            MemoryStream stream = CreateWorkbook(true, cellFormats, sheetData =>
            {
                sheetData.Append(new Row(
                    new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue("1") },
                    new Cell { CellReference = "B1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("inline")) })
                { RowIndex = 1 });
            });

            ToxySpreadsheet ss = ParserFactory.CreateSpreadsheet(new ParserContext(stream)).Parse();
            // day 1 is 1904-01-02 in the 1904 date system, it would be 1900-01-01 in the 1900 date system
            ClassicAssert.AreEqual("01-02-04", ss.Tables[0][0][0].ToString());
            ClassicAssert.AreEqual("inline", ss.Tables[0][0][1].ToString());
        }

        private static MemoryStream CreateWorkbook(bool date1904, CellFormats cellFormats, Action<SheetData> fillSheet)
        {
            MemoryStream stream = new MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                if (date1904)
                {
                    workbookPart.Workbook.Append(new WorkbookProperties { Date1904 = true });
                }
                if (cellFormats != null)
                {
                    WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    stylesPart.Stylesheet = new Stylesheet(cellFormats);
                }

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                fillSheet(sheetData);
                worksheetPart.Worksheet = new Worksheet(sheetData);

                workbookPart.Workbook.Append(new Sheets(
                    new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
            }
            stream.Position = 0;
            return stream;
        }
    }
}

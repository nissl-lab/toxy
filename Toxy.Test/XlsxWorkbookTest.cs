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
    /// Tests for .xlsx features read by ExcelDataReader which are not covered by the files in testdata, the workbooks are created in memory.
    /// </summary>
    [TestFixture]
    public class XlsxWorkbookTest
    {
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

        [Test]
        public void TestStreamStaysOpen()
        {
            MemoryStream stream = CreateWorkbook(false, null, sheetData =>
                sheetData.Append(new Row(new Cell { CellReference = "A1", CellValue = new CellValue("1") }) { RowIndex = 1 }));

            ParserContext context = new ParserContext(stream);
            ParserFactory.CreateSpreadsheet(context).Parse();
            ParserFactory.CreateSpreadsheet(context).Parse(0);

            // a stream passed in by the caller is not closed by the parser
            ClassicAssert.IsTrue(stream.CanRead);
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

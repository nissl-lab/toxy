using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    public class ExcelParserBaseTest
    {
        public void BaseTestShowCalculatedResult(string filename)
        { 
            
        }
        public void BaseTestFooter(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.IsNull(ss.Tables[0].PageFooter);

            parser.Context.Properties.Add("ExtractSheetFooter", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            Assert.IsNotNull(ss2.Tables[0].PageFooter);
            Assert.AreEqual("testdoc|test phrase|",ss2.Tables[0].PageFooter);
        }
        public void BaseTestWithoutHeader(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.IsNull(ss.Tables[0].PageHeader);

            ToxySpreadsheet ss2 = parser.Parse();
            Assert.AreEqual(0, ss.Tables[0].ColumnHeaders.Cells.Count);
            Assert.AreEqual(9, ss.Tables[0].Rows.Count);
        }
        public void BaseTestHeader(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.IsNull(ss.Tables[0].PageHeader);

            parser.Context.Properties.Add("ExtractSheetHeader", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            Assert.IsNotNull(ss2.Tables[0].PageHeader);
            Assert.AreEqual("|testdoc|test phrase", ss2.Tables[0].PageHeader);
        }
        public void BaseTestFillBlankCells(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.AreEqual(1, ss.Tables[0].Rows[0].Cells.Count);
            Assert.AreEqual(0, ss.Tables[0].Rows[1].Cells.Count);
            Assert.AreEqual(2, ss.Tables[0].Rows[2].Cells.Count);
            Assert.AreEqual(2, ss.Tables[0].Rows[3].Cells.Count);
            Assert.AreEqual(2, ss.Tables[0].Rows[4].Cells.Count);

            parser.Context.Properties.Add("FillBlankCells", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            
            Assert.AreEqual(3, ss2.Tables[0].Rows[0].Cells.Count);
            Assert.AreEqual(3, ss2.Tables[0].Rows[1].Cells.Count);
            Assert.AreEqual(3, ss2.Tables[0].Rows[2].Cells.Count);
            Assert.AreEqual(3, ss2.Tables[0].Rows[3].Cells.Count);
            Assert.AreEqual(3, ss2.Tables[0].Rows[4].Cells.Count);
        }

        public void BaseTestExcelContent(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.AreEqual(3, ss.Tables.Count);
            Assert.AreEqual("Sheet1", ss.Tables[0].Name);
            Assert.AreEqual("Sheet2", ss.Tables[1].Name);
            Assert.AreEqual("Sheet3", ss.Tables[2].Name);
            Assert.AreEqual(5, ss.Tables[0].Rows.Count);
            Assert.AreEqual(0, ss.Tables[1].Rows.Count);
            Assert.AreEqual(0, ss.Tables[2].Rows.Count);

            ToxyTable table = ss.Tables[0];
            Assert.AreEqual(1, table.Rows[0].RowIndex);
            Assert.AreEqual(2, table.Rows[1].RowIndex);
            Assert.AreEqual(3, table.Rows[2].RowIndex);
            Assert.AreEqual(4, table.Rows[3].RowIndex);
            Assert.AreEqual(5, table.Rows[4].RowIndex);


            Assert.AreEqual(1, table.Rows[0].Cells.Count);
            Assert.AreEqual(0, table.Rows[1].Cells.Count);
            Assert.AreEqual(2, table.Rows[2].Cells.Count);
            Assert.AreEqual(2, table.Rows[3].Cells.Count);
            Assert.AreEqual(2, table.Rows[4].Cells.Count);
            Assert.AreEqual("Employee Info", table.Rows[0].Cells[0].ToString());
            Assert.AreEqual(1, table.Rows[0].Cells[0].CellIndex);
            Assert.AreEqual("Last name:", table.Rows[2].Cells[0].ToString());
            Assert.AreEqual(1, table.Rows[2].Cells[0].CellIndex);
            Assert.AreEqual("lastName", table.Rows[2].Cells[1].ToString());
            Assert.AreEqual(2, table.Rows[2].Cells[1].CellIndex);
            Assert.AreEqual("First name:", table.Rows[3].Cells[0].ToString());
            Assert.AreEqual("firstName", table.Rows[3].Cells[1].ToString());
            Assert.AreEqual("SSN:", table.Rows[4].Cells[0].ToString());
            Assert.AreEqual("ssn", table.Rows[4].Cells[1].ToString());
        }

        public void BaseTestExcelComment(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.AreEqual(3, ss.Tables.Count);
            Assert.AreEqual(3, ss.Tables[0].Rows.Count);
            Assert.AreEqual(0, ss.Tables[0].Rows[1].Cells[0].CellIndex);
            Assert.AreEqual(2, ss.Tables[0].Rows[1].RowIndex);
            //TODO: fix this comment without cell value
            Assert.AreEqual("comment top row1 (index0)\n", ss.Tables[0].Rows[0].Cells[0].Comment);
            Assert.AreEqual("comment top row3 (index2)\n", ss.Tables[0].Rows[1].Cells[0].Comment);
            Assert.AreEqual("comment top row4 (index3)\n", ss.Tables[0].Rows[2].Cells[0].Comment);
        }

        public void BaseTestExcelFormatedString(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            Assert.AreEqual(13, ss.Tables[0].Rows.Count);
            Assert.AreEqual("Dates, all 24th November 2006", ss.Tables[0].Rows[0].Cells[0].ToString());
            Assert.AreEqual("24/11/2006", ss.Tables[0].Rows[1].Cells[1].ToString());
            Assert.AreEqual("2006/11/24", ss.Tables[0].Rows[2].Cells[1].ToString());
            Assert.AreEqual("2006-11-24", ss.Tables[0].Rows[3].Cells[1].ToString());
            Assert.AreEqual("06/11/24", ss.Tables[0].Rows[4].Cells[1].ToString());
            Assert.AreEqual("24/11/06", ss.Tables[0].Rows[5].Cells[1].ToString());
            Assert.AreEqual("24-11-06", ss.Tables[0].Rows[6].Cells[1].ToString());

            Assert.AreEqual("10.52", ss.Tables[0].Rows[9].Cells[1].ToString());
            Assert.AreEqual("10.520", ss.Tables[0].Rows[10].Cells[1].ToString());
            Assert.AreEqual("10.5", ss.Tables[0].Rows[11].Cells[1].ToString());
            Assert.AreEqual("£10.52", ss.Tables[0].Rows[12].Cells[1].ToString());
            
        }
    }
}

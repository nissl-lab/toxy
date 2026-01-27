using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace Toxy.Test
{
    public class ExcelParserBaseTest
    {
        public void BaseTestShowCalculatedResult(string filename)
        { 
            
        }
        public void BaseTestExtractSheetFooter(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.IsNull(ss.Tables[0].PageFooter);

            parser.Context.Properties.Add("ExtractSheetFooter", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            ClassicAssert.IsNotNull(ss2.Tables[0].PageFooter);
            ClassicAssert.AreEqual("testdoc|test phrase|",ss2.Tables[0].PageFooter);
        }
        public void BaseTestWithoutHeader(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.IsNull(ss.Tables[0].PageHeader);

            ClassicAssert.AreEqual(0, ss.Tables[0].HeaderRows.Count);
            ClassicAssert.AreEqual(9, ss.Tables[0].Length);
        }
        public void BaseTestWithHeaderRow(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            parser.Context.Properties.Add("HasHeader", "1");
            ToxySpreadsheet ss = parser.Parse();

            ClassicAssert.AreEqual(1, ss.Tables[0].HeaderRows.Count);
            ClassicAssert.AreEqual("A", ss.Tables[0].HeaderRows[0][0].Value);
            ClassicAssert.AreEqual("B", ss.Tables[0].HeaderRows[0][1].Value);
            ClassicAssert.AreEqual("C", ss.Tables[0].HeaderRows[0][2].Value);
            ClassicAssert.AreEqual("D", ss.Tables[0].HeaderRows[0][3].Value);
            ClassicAssert.AreEqual(3, ss.Tables[0].Length);
            ClassicAssert.AreEqual("1", ss.Tables[0][0][0].Value);
            ClassicAssert.AreEqual("2", ss.Tables[0][0][1].Value);
            ClassicAssert.AreEqual("3", ss.Tables[0][0][2].Value);
            ClassicAssert.AreEqual("4", ss.Tables[0][0][3].Value);

            ClassicAssert.AreEqual("A1", ss.Tables[0][1][0].Value);
            ClassicAssert.AreEqual("A2", ss.Tables[0][1][1].Value);
            ClassicAssert.AreEqual("A3", ss.Tables[0][1][2].Value);
            ClassicAssert.AreEqual("A4", ss.Tables[0][1][3].Value);

            ClassicAssert.AreEqual("B1", ss.Tables[0][2][0].Value);
            ClassicAssert.AreEqual("B2", ss.Tables[0][2][1].Value);
            ClassicAssert.AreEqual("B3", ss.Tables[0][2][2].Value);
            ClassicAssert.AreEqual("B4", ss.Tables[0][2][3].Value);
        }

        public void BaseTestSlicedTableAndRow(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            
            var table1_2 = ss[1..3];    //it only has 3 tables
            ClassicAssert.AreEqual(2, table1_2.Length);

            ToxyTable table0 = ss.Tables[0];
            var slicedrows= table0[1..4];
            ClassicAssert.AreEqual("Last name:", slicedrows[1][0].ToString());
            ClassicAssert.AreEqual("lastName", slicedrows[1][1].ToString());
            ClassicAssert.AreEqual("First name:", slicedrows[2][0].ToString());
            ClassicAssert.AreEqual("firstName", slicedrows[2][1].ToString());
        }

        public void BaseTestExtractSheetHeader(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.IsNull(ss.Tables[0].PageHeader);

            parser.Context.Properties.Add("ExtractSheetHeader", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            ClassicAssert.IsNotNull(ss2.Tables[0].PageHeader);
            ClassicAssert.AreEqual("|testdoc|test phrase", ss2.Tables[0].PageHeader);
        }
        public void BaseTestFillBlankCells(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.AreEqual(1, ss.Tables[0][0].Length);
            ClassicAssert.AreEqual(0, ss.Tables[0][1].Length);
            ClassicAssert.AreEqual(2, ss.Tables[0][2].Length);
            ClassicAssert.AreEqual(2, ss.Tables[0][3].Length);
            ClassicAssert.AreEqual(2, ss.Tables[0][4].Length);

            parser.Context.Properties.Add("FillBlankCells", "1");
            ToxySpreadsheet ss2 = parser.Parse();
            
            ClassicAssert.AreEqual(3, ss2.Tables[0][0].Length);
            ClassicAssert.AreEqual(3, ss2.Tables[0][1].Length);
            ClassicAssert.AreEqual(3, ss2.Tables[0][2].Length);
            ClassicAssert.AreEqual(3, ss2.Tables[0][3].Length);
            ClassicAssert.AreEqual(3, ss2.Tables[0][4].Length);
        }

        public void BaseTestExcelContent(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.AreEqual(3, ss.Tables.Count);
            ClassicAssert.AreEqual("Sheet1", ss.Tables[0].Name);
            ClassicAssert.AreEqual("Sheet2", ss.Tables[1].Name);
            ClassicAssert.AreEqual("Sheet3", ss.Tables[2].Name);
            ClassicAssert.AreEqual(5, ss.Tables[0].Length);
            ClassicAssert.AreEqual(0, ss.Tables[1].Length);
            ClassicAssert.AreEqual(0, ss.Tables[2].Length);

            ToxyTable table = ss.Tables[0];
            ClassicAssert.AreEqual(1, table[0].RowIndex);
            ClassicAssert.AreEqual(2, table[1].RowIndex);
            ClassicAssert.AreEqual(3, table[2].RowIndex);
            ClassicAssert.AreEqual(4, table[3].RowIndex);
            ClassicAssert.AreEqual(5, table[4].RowIndex);


            ClassicAssert.AreEqual(1, table[0].Length);
            ClassicAssert.AreEqual(0, table[1].Length);
            ClassicAssert.AreEqual(2, table[2].Length);
            ClassicAssert.AreEqual(2, table[3].Length);
            ClassicAssert.AreEqual(2, table[4].Length);
            ClassicAssert.AreEqual("Employee Info", table[0][0].ToString());
            ClassicAssert.AreEqual(1, table[0][0].CellIndex);
            ClassicAssert.AreEqual("Last name:", table[2][0].ToString());
            ClassicAssert.AreEqual(1, table[2][0].CellIndex);
            ClassicAssert.AreEqual("lastName", table[2][1].ToString());
            ClassicAssert.AreEqual(2, table[2][1].CellIndex);
            ClassicAssert.AreEqual("First name:", table[3][0].ToString());
            ClassicAssert.AreEqual("firstName", table[3][1].ToString());
            ClassicAssert.AreEqual("SSN:", table[4][0].ToString());
            ClassicAssert.AreEqual("ssn", table[4][1].ToString());
        }

        public void BaseTestExcelComment(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.AreEqual(3, ss.Tables.Count);
            ClassicAssert.AreEqual(3, ss.Tables[0].Length);
            ClassicAssert.AreEqual(0, ss.Tables[0][1][0].CellIndex);
            ClassicAssert.AreEqual(2, ss.Tables[0][1].RowIndex);
            //TODO: fix this comment without cell value
            ClassicAssert.AreEqual("comment top row1 (index0)\n", ss.Tables[0][0][0].Comment);
            ClassicAssert.AreEqual("comment top row3 (index2)\n", ss.Tables[0][1][0].Comment);
            ClassicAssert.AreEqual("comment top row4 (index3)\n", ss.Tables[0][2][0].Comment);
        }

        public void BaseTestExcelFormatedString(string filename)
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath(filename));
            ISpreadsheetParser parser = ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss = parser.Parse();
            ClassicAssert.AreEqual(13, ss.Tables[0].Length);
            ClassicAssert.AreEqual("Dates, all 24th November 2006", ss.Tables[0][0][0].ToString());
            ClassicAssert.AreEqual("24/11/2006", ss.Tables[0][1][1].ToString());
            ClassicAssert.AreEqual("2006/11/24", ss.Tables[0][2][1].ToString());
            ClassicAssert.AreEqual("2006-11-24", ss.Tables[0][3][1].ToString());
            ClassicAssert.AreEqual("06/11/24", ss.Tables[0][4][1].ToString());
            ClassicAssert.AreEqual("24/11/06", ss.Tables[0][5][1].ToString());
            ClassicAssert.AreEqual("24-11-06", ss.Tables[0][6][1].ToString());

            ClassicAssert.AreEqual("10.52", ss.Tables[0][9][1].ToString());
            ClassicAssert.AreEqual("10.520", ss.Tables[0][10][1].ToString());
            ClassicAssert.AreEqual("10.5", ss.Tables[0][11][1].ToString());
            ClassicAssert.AreEqual("£10.52", ss.Tables[0][12][1].ToString());
            
        }
    }
}

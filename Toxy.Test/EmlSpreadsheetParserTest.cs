using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace Toxy.Test
{
    [TestFixture]
    public class EmlSpreadsheetParserTest
    {
        [Test]
        public void EML_MarkdownTable_ParsesCorrectly()
        {
            string path = TestDataSample.GetEmailPath("tables.eml");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();

            ClassicAssert.AreEqual(2, spreadsheet.Tables.Count);

            // First table: Product/Sales/Region
            var table1 = spreadsheet.Tables[0];
            ClassicAssert.IsTrue(table1.HasHeader);
            ClassicAssert.AreEqual(1, table1.HeaderRows.Count);
            ClassicAssert.AreEqual("Product", table1.HeaderRows[0].Cells[0].Value);
            ClassicAssert.AreEqual("Sales", table1.HeaderRows[0].Cells[1].Value);
            ClassicAssert.AreEqual("Region", table1.HeaderRows[0].Cells[2].Value);
            ClassicAssert.AreEqual(2, table1.Rows.Count);
            ClassicAssert.AreEqual("Widget", table1.Rows[0].Cells[0].Value);
            ClassicAssert.AreEqual("1200", table1.Rows[0].Cells[1].Value);
            ClassicAssert.AreEqual("North", table1.Rows[0].Cells[2].Value);
            ClassicAssert.AreEqual("Gadget", table1.Rows[1].Cells[0].Value);

            // Second table: Region/Q1/Q2/Q3
            var table2 = spreadsheet.Tables[1];
            ClassicAssert.IsTrue(table2.HasHeader);
            ClassicAssert.AreEqual("Region", table2.HeaderRows[0].Cells[0].Value);
            ClassicAssert.AreEqual(3, table2.Rows.Count);
            ClassicAssert.AreEqual("North", table2.Rows[0].Cells[0].Value);
            ClassicAssert.AreEqual("100", table2.Rows[0].Cells[1].Value);
        }

        [Test]
        public void EML_HtmlTable_ParsesCorrectly()
        {
            string path = TestDataSample.GetEmailPath("html-tables.eml");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();

            ClassicAssert.AreEqual(2, spreadsheet.Tables.Count);

            var table1 = spreadsheet.Tables[0];
            ClassicAssert.AreEqual("products", table1.Name);
            ClassicAssert.AreEqual(3, table1.Rows.Count);
            ClassicAssert.AreEqual("Product", table1.Rows[0].Cells[0].Value);
            ClassicAssert.AreEqual("Widget", table1.Rows[1].Cells[0].Value);

            var table2 = spreadsheet.Tables[1];
            ClassicAssert.AreEqual("regions", table2.Name);
            ClassicAssert.AreEqual(3, table2.Rows.Count);
        }

        [Test]
        public void EML_NoMarkdownTables_HtmlTablesExtracted()
        {
            // test.eml has an HTML body — HTML tables should be extracted
            string path = TestDataSample.GetEmailPath("test.eml");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();

            // test.eml has HTML body with at least one table
            ClassicAssert.IsNotNull(spreadsheet);
            ClassicAssert.GreaterOrEqual(spreadsheet.Tables.Count, 1);
        }
    }
}

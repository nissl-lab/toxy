using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

namespace Toxy.Test
{
    [TestFixture]
    public class TestHtmlSpreadsheetParser
    {
        [Test]
        public void TestParse()
        {
            string path = TestDataSample.GetHtmlPath("tables.html");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();
            
            ClassicAssert.AreEqual(2, spreadsheet.Tables.Count);
            ClassicAssert.AreEqual("users", spreadsheet.Tables[0].Name);
            ClassicAssert.AreEqual("products", spreadsheet.Tables[1].Name);
        }

        [Test]
        public void TestHtmlSpreadsheetParserFirstTable()
        {
            string path = TestDataSample.GetHtmlPath("tables.html");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var table = parser.Parse(0);
            
            ClassicAssert.AreEqual("users", table.Name);
            ClassicAssert.AreEqual(3, table.Rows.Count);
            ClassicAssert.AreEqual(3, table.Rows[0].Cells.Count);
            ClassicAssert.AreEqual("Name", table.Rows[0].Cells[0].Value);
            ClassicAssert.AreEqual("Age", table.Rows[0].Cells[1].Value);
            ClassicAssert.AreEqual("City", table.Rows[0].Cells[2].Value);
            ClassicAssert.AreEqual("John", table.Rows[1].Cells[0].Value);
            ClassicAssert.AreEqual("30", table.Rows[1].Cells[1].Value);
            ClassicAssert.AreEqual("Jane", table.Rows[2].Cells[0].Value);
        }

        [Test]
        public void TestHtmlSpreadsheetParserNoTable()
        {
            string path = TestDataSample.GetHtmlPath("mshome.html");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();
            
            ClassicAssert.AreEqual(0, spreadsheet.Tables.Count);
        }

        [Test]
        public void TestHtmlSpreadsheetParserDirect()
        {
            var parser = new Toxy.Parsers.HtmlSpreadsheetParser(new ParserContext(TestDataSample.GetHtmlPath("tables.html")));
            var spreadsheet = parser.Parse();
            
            ClassicAssert.AreEqual(2, spreadsheet.Tables.Count);
            ClassicAssert.AreEqual("users", spreadsheet.Tables[0].Name);
        }

        [Test]
        public void TestHtmlSpreadsheetParserInnerTextOnly()
        {
            string html = @"<html><body><table><tr><td><b>Bold</b> and <i>italic</i></td></tr></table></body></html>";
            using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
            var parser = new Toxy.Parsers.HtmlSpreadsheetParser(new ParserContext(stream));
            var spreadsheet = parser.Parse();
            
            ClassicAssert.AreEqual("Bold and italic", spreadsheet.Tables[0].Rows[0].Cells[0].Value);
        }
    }
}

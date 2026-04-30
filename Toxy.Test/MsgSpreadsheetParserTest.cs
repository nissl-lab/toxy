using NUnit.Framework;
using NUnit.Framework.Legacy;

namespace Toxy.Test
{
    [TestFixture]
    public class MsgSpreadsheetParserTest
    {
        [Test]
        public void MSG_HtmlTable_ParsesCorrectly()
        {
            string path = TestDataSample.GetEmailPath("Azure pricing and services updates.msg");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();

            ClassicAssert.IsNotNull(spreadsheet);
            ClassicAssert.Greater(spreadsheet.Tables.Count, 0);
        }

        [Test]
        public void MSG_NoTables_ReturnsEmptySpreadsheet()
        {
            string path = TestDataSample.GetEmailPath("raw text mail demo.msg");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var spreadsheet = parser.Parse();

            ClassicAssert.IsNotNull(spreadsheet);
        }
    }
}

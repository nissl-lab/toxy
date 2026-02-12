using NUnit.Framework;

namespace Toxy.Test
{
    /// <summary>
    /// Summary description for ExcelParserTest
    /// </summary>
    [TestFixture]
    public class Excel2003ParserTest:ExcelParserBaseTest
    {
        public Excel2003ParserTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        [Test]
        public void TestExtractFooter()
        {
            BaseTestExtractSheetFooter("45538_classic_Footer.xls");
        }
        [Test]
        public void TestExtractHeader()
        {
            BaseTestExtractSheetHeader("45538_classic_Header.xls");
        }
        [Test]
        public void TestFillBlankCells()
        {
            base.BaseTestFillBlankCells("Employee.xls");
        }

        [Test]
        public void TestExcelWithFormats()
        {
            BaseTestExcelFormatedString("Formatting.xls");
        }
        [Test]
        public void TestSlicedRow()
        {
            BaseTestSlicedRow("Employee.xls");
        }
        [Test]
        public void TestSlicedCell()
        {
            BaseTestSlicedCell("Employee.xls");
        }
        [Test]
        public void TestSlicedTable()
        {
            BaseTestSlicedTable("Employee.xls");
        }
        [Test]
        public void TestExcelParserSimple()
        {
            base.BaseTestExcelContent("Employee.xls");
        }
        [Test]
        public void TestExcelWithComments()
        {
            base.BaseTestExcelComment("comments.xls");
        }
    }
}

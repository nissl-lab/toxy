using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Toxy.Test
{
    /// <summary>
    /// Summary description for ExcelParserTest
    /// </summary>
    [TestClass]
    public class Excel2003ParserTest:ExcelParserBaseTest
    {
        public Excel2003ParserTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        [TestMethod]
        public void TestExtractFooter()
        {
            BaseTestFooter("45538_classic_Footer.xls");
        }
        [TestMethod]
        public void TestExtractHeader()
        {
            BaseTestHeader("45538_classic_Header.xls");
        }
        [TestMethod]
        public void TestFillBlankCells()
        {
            base.BaseTestFillBlankCells("Employee.xls");
        }

        [TestMethod]
        public void TestExcelWithFormats()
        {
            BaseTestExcelFormatedString("Formatting.xls");
        }

        [TestMethod]
        public void TestExcelParserSimple()
        {
            base.BaseTestExcelContent("Employee.xls");
        }
        [TestMethod]
        public void TestExcelWithComments()
        {
            base.BaseTestExcelComment("comments.xls");
        }
    }
}

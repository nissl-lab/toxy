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
    public class Excel2007ParserTest:ExcelParserBaseTest
    {
        public Excel2007ParserTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [TestMethod]
        public void TestExtractFooter()
        {
            BaseTestFooter("45540_classic_Footer.xlsx");
        }
        [TestMethod]
        public void TestExtractHeader()
        {
            BaseTestHeader("45540_classic_Header.xlsx");
        }
        [TestMethod]
        public void TestExtractWithoutHeader()
        {
            BaseTestWithoutHeader("WithVariousData.xlsx");
        }
        [TestMethod]
        public void TestExcelWithFormats()
        {
            BaseTestExcelFormatedString("Formatting.xlsx");
        }
        
        [TestMethod]
        public void TestExcelWithComments()
        {
            base.BaseTestExcelComment("comments.xlsx");
        }
    }
}

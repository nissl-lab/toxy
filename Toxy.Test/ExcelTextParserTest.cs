using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class ExcelTextParserTest
    {
        [Test]
        public void TestExcel2003TextParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("Employee.xls"));
            ITextParser parser = ParserFactory.CreateText(context);
            string result= parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IndexOf("Last name")>0);
            Assert.IsTrue(result.IndexOf("First name") > 0);
        }
        [Test]
        public void TestExcel2007TextParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IndexOf("Foo") > 0);
            Assert.IsTrue(result.IndexOf("Bar") > 0);
            Assert.IsTrue(result.IndexOf("a really long cell") > 0);

            Assert.IsTrue(result.IndexOf("have a header") > 0);
            Assert.IsTrue(result.IndexOf("have a footer") > 0);
            Assert.IsTrue(result.IndexOf("This is the header") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithoutComment()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeComments","0");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IndexOf("Comment by") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithoutSheetNames()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeSheetNames", "0");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IndexOf("Sheet1") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithHeaderFooter()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeHeaderFooter", "1");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.IndexOf("This is the header") > 0);
        }
    }
}

using NUnit.Framework;
using NUnit.Framework.Legacy;

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
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.IndexOf("Last name")>0);
            ClassicAssert.IsTrue(result.IndexOf("First name") > 0);
        }
        [Test]
        public void TestExcel2007TextParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.IndexOf("Foo") > 0);
            ClassicAssert.IsTrue(result.IndexOf("Bar") > 0);
            ClassicAssert.IsTrue(result.IndexOf("a really long cell") > 0);

            ClassicAssert.IsTrue(result.IndexOf("have a header") > 0);
            ClassicAssert.IsTrue(result.IndexOf("have a footer") > 0);
            ClassicAssert.IsTrue(result.IndexOf("This is the header") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithoutComment()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeComments","0");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.IndexOf("Comment by") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithoutSheetNames()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeSheetNames", "0");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.IndexOf("Sheet1") < 0);
        }
        [Test]
        public void TestExcel2007TextParserWithHeaderFooter()
        {
            ParserContext context = new ParserContext(TestDataSample.GetExcelPath("WithVariousData.xlsx"));
            context.Properties.Add("IncludeHeaderFooter", "1");
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.IndexOf("This is the header") > 0);
        }
    }
}

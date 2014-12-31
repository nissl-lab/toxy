using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class OLE2MetadataParserTest
    {
        [Test]
        public void TestWord()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestEditTime.doc"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(19, x.Count);

            path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestChineseProperties.doc"));
            context = new ParserContext(path);
            parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            x = parser.Parse();
            Assert.AreEqual(19, x.Count);
            Assert.AreEqual("雅虎網站分類", x.Get("Comments").Value);
            Assert.AreEqual("參考資料", x.Get("Title").Value);
        }
        [Test]
        public void TestExcelFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetExcelPath("comments.xls"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(12, x.Count);
            Assert.AreEqual("Microsoft Excel", x.Get("ApplicationName").Value);
        }
        [Test]
        public void TestPowerPoint()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("Test_Humor-Generation.ppt"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(11, x.Count);
            Assert.AreEqual("Funny Factory", x.Get("Title").Value);
        }
        [Test]
        public void TestCorelDrawFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestCorel.shw"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(10, x.Count);
            Assert.AreEqual("thorsteb", x.Get("Author").Value);
            Assert.AreEqual("thorsteb", x.Get("LastAuthor").Value);
        }
        [Test]
        public void TestSolidWorksFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestSolidWorks.sldprt"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            Assert.AreEqual(14, x.Count);
            Assert.AreEqual("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", x.Get("ClassID").Value);
            Assert.AreEqual("scj", x.Get("LastAuthor").Value);
        }
    }
}

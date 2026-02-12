using NUnit.Framework;
using NUnit.Framework.Legacy;
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
        [Ignore(".doc extraction is not implemented")]
        public void TestWord()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestEditTime.doc"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            ClassicAssert.AreEqual(18, x.Count);

            path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestChineseProperties.doc"));
            context = new ParserContext(path);
            parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            x = parser.Parse();
            ClassicAssert.AreEqual(18, x.Count);
            ClassicAssert.AreEqual("雅虎網站分類", x.Get("Comments").Value);
            ClassicAssert.AreEqual("參考資料", x.Get("Title").Value);
        }
        [Test]
        public void TestExcelFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetExcelPath("comments.xls"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            ClassicAssert.AreEqual(8, x.Count);
            ClassicAssert.AreEqual("Microsoft Excel", x.Get("ApplicationName").Value);
        }
        [Test]
        public void TestPowerPoint()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("Test_Humor-Generation.ppt"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            ClassicAssert.AreEqual(8, x.Count);
            ClassicAssert.AreEqual("Funny Factory", x.Get("Title").Value);
        }
        [Test]
        public void TestCorelDrawFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestCorel.shw"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            ClassicAssert.AreEqual(6, x.Count);
            ClassicAssert.AreEqual("thorsteb", x.Get("Author").Value);
            ClassicAssert.AreEqual("thorsteb", x.Get("LastAuthor").Value);
        }
        [Test]
        public void TestSolidWorksFile()
        {
            string path = Path.GetFullPath(TestDataSample.GetOLE2Path("TestSolidWorks.sldprt"));
            ParserContext context = new ParserContext(path);
            IMetadataParser parser = (IMetadataParser)ParserFactory.CreateMetadata(context);
            ToxyMetadata x = parser.Parse();
            ClassicAssert.AreEqual(10, x.Count);
            ClassicAssert.AreEqual("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}", x.Get("ClassID").Value);
            ClassicAssert.AreEqual("scj", x.Get("LastAuthor").Value);
        }
    }
}

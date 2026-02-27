using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;

namespace Toxy.Test
{
    [TestFixture]
    public class ZipParserTest
    {
        [Test]
        public void TestParseDirectoryFromZip()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFilePath("toxy.zip", null));
            ITextParser parser = ParserFactory.CreateText(context);
            string list = parser.Parse();
            ClassicAssert.IsNotNull(list);
            string[] lines = list.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual(68, lines.Length);
        }
        [Test]
        public void TestStreamFromZip()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("toxy.zip",null));
            ITextParser parser = ParserFactory.CreateText(context);
            string list = parser.Parse();
        }
        [Test]
        public void TestEncryptedZip()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFilePath("password.zip", null));
            ITextParser parser = ParserFactory.CreateText(context);
            ClassicAssert.Throws<InvalidOperationException>(() => parser.Parse());
        }
    }
}

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
                ParserContext context = new ParserContext(TestDataSample.GetFilePath("toxy.zip",null));
                ITextParser parser = ParserFactory.CreateText(context);
                string list = parser.Parse();
                ClassicAssert.IsNotNull(list);
                string[] lines = list.Split(new string[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
                ClassicAssert.AreEqual(68, lines.Length);
            }
    }
}

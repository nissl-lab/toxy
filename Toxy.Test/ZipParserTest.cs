using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                Assert.IsNotNull(list);
                string[] lines = list.Split(new string[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
                Assert.AreEqual(68, lines.Length);
            }
    }
}

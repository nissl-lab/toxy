using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class Powerpoint2007TextParserTest
    {
        [Test]
        public void ReadTextBasicTest()
        {
            string path = Path.GetFullPath(TestDataSample.GetPowerpointPath("testPPT.pptx"));
            ParserContext context = new ParserContext(path);
            ITextParser parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
            Assert.IsNotNullOrEmpty(result);
            string[] texts = result.Split(new string[]{"\r\n"}, StringSplitOptions.RemoveEmptyEntries);
            Assert.AreEqual(14, texts.Length);
            Assert.AreEqual("Attachment Test", texts[0]);
            Assert.AreEqual("Rajiv", texts[1]);
            Assert.AreEqual("Different words to test against", texts[4]);
            Assert.AreEqual("Hello", texts[7]);
        }
    }
}
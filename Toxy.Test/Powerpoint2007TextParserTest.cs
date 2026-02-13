using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;

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
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsNotEmpty(result);
            string[] texts = result.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual(14, texts.Length);
            ClassicAssert.AreEqual("Attachment Test", texts[0]);
            ClassicAssert.AreEqual("Rajiv", texts[1]);
            ClassicAssert.AreEqual("Different words to test against", texts[4]);
            ClassicAssert.AreEqual("Hello", texts[7]);
        }
    }
}
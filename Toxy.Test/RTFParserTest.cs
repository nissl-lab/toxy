using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Toxy.Parsers;

namespace Toxy.Test
{
    [TestFixture]
    public class RTFParserTest
    {
        [Test]
        public void TestReadRTF_FormattedText()
        {
            string path = TestDataSample.GetRTFPath("Formated text.rtf");
            var parser = new RTFTextParser(new ParserContext(path));
            string result = parser.Parse();
            string[] lines = result.Replace("\r\n", "\n").Split('\n');
            Assert.AreEqual(lines.Length, 11);
            Assert.AreEqual("11111111111", lines[0]);
            Assert.AreEqual("22222222222", lines[1]);
            Assert.AreEqual("张三李四王五", lines[2]);
            Assert.AreEqual("RTF Sample , Author : yuans , contact : yyf9989@hotmail.com , site : http://www.cnblogs.com/xdesigner .", lines[7]);
        }
        [Test]
        public void TestReadRTF_Html()
        {
            string path = TestDataSample.GetRTFPath("htmlrtf2.rtf");
            var parser = new RTFTextParser(new ParserContext(path));
            string result = parser.Parse();
            Assert.IsNotNullOrEmpty(result);
        }
    }
}

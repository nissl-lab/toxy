using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;
using ToxyFramework.Parsers;

namespace Toxy.Test
{
    [TestFixture]
    public class RTFParserTest
    {
        [Test]
        public void TestReadRTF_FormattedText()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string path = TestDataSample.GetRTFPath("Formated text.rtf");
            var context = new ParserContext(path);
            context.Encoding = Encoding.GetEncoding("windows-1252");
            var parser = new RTFTextParser(context);
            string result = parser.Parse();
            string[] lines = result.Replace("\r\n", "\n").Split('\n');
            Assert.AreEqual(lines.Length, 10);
            Assert.AreEqual("11111111111", lines[0]);
            Assert.AreEqual("22222222222", lines[1]);
            Assert.AreEqual("RTF Sample , Author : yuans , contact : yyf9989@hotmail.com , site : http://www.cnblogs.com/xdesigner .", lines[7]);

            Assert.AreEqual("张三李四王五", lines[2]); //encoding issue
        }
        [Test]
        public void TestReadRTF_Html()
        {
            string path = TestDataSample.GetRTFPath("htmlrtf2.rtf");
            var parser = new RTFTextParser(new ParserContext(path));
            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsTrue(result.Contains("Beste CMMA van Spelde,"));
        }
    }
}

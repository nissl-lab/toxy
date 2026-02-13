using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string path = TestDataSample.GetRTFPath("Formated text.rtf");
            var context = new ParserContext(path);
            context.Encoding = Encoding.GetEncoding("windows-1252");
            var parser = new RTFTextParser(context);
            string result = parser.Parse();
            ClassicAssert.IsTrue(result.Contains("11111111111"));
            ClassicAssert.IsTrue(result.Contains("22222222222"));
            ClassicAssert.IsTrue(result.Contains("RTF Sample , Author : yuans , contact : yyf9989@hotmail.com , site : http://www.cnblogs.com/xdesigner ."));

            ClassicAssert.IsTrue(result.Contains("张三李四王五")); //encoding issue
        }
        [Test]
        public void TestReadRTF_Html()
        {
            string path = TestDataSample.GetRTFPath("htmlrtf2.rtf");
            var parser = new RTFTextParser(new ParserContext(path));
            string result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.Contains("Beste CMMA van Spelde,"));
        }
    }
}

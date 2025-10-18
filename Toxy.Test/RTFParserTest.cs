using NUnit.Framework;
using NUnit.Framework.Legacy;
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
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            string path = TestDataSample.GetRTFPath("Formated text.rtf");
            var context = new ParserContext(path);
            context.Encoding = Encoding.GetEncoding("windows-1252");
            var parser = new RTFTextParser(context);
            string result = parser.Parse();
            string[] lines = result.Replace("\r\n", "\n").Split('\n');
            ClassicAssert.AreEqual(lines.Length, 10);
            ClassicAssert.AreEqual("11111111111", lines[0]);
            ClassicAssert.AreEqual("22222222222", lines[1]);
            ClassicAssert.AreEqual("RTF Sample , Author : yuans , contact : yyf9989@hotmail.com , site : http://www.cnblogs.com/xdesigner .", lines[7]);

            ClassicAssert.AreEqual("张三李四王五", lines[2]); //encoding issue
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

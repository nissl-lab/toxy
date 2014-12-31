using System;
using NUnit.Framework;
using Toxy.Parsers;

namespace Toxy.Test
{
    [TestFixture]
    public class PlainTextParserTest
    {
        [Test]
        public void TestReadWholeText()
        {
            string path = TestDataSample.GetTextPath("utf8.txt");

            ParserContext context=new ParserContext(path);
            ITextParser parser = ParserFactory.CreateText(context);
            string text= parser.Parse();
            Assert.AreEqual("hello world\r\na2\r\na3\r\nbbb4\r\n", text);
        }

        [Test]
        public void TestParseLineEvent()
        {
            string path = TestDataSample.GetTextPath("utf8.txt");
            ParserContext context = new ParserContext(path);
            PlainTextParser parser = (PlainTextParser)ParserFactory.CreateText(context);
            parser.ParseLine += (sender, args) => 
            {
                if (args.LineNumber == 0)
                {
                    Assert.AreEqual("hello world", args.Text);
                }
                else if(args.LineNumber==1)
                {
                    Assert.AreEqual("a2", args.Text);
                }
                else if (args.LineNumber == 2)
                {
                    Assert.AreEqual("a3", args.Text);
                }
                else if (args.LineNumber == 3)
                {
                    Assert.AreEqual("bbb4", args.Text);
                }
            };
            string text = parser.Parse();
        }

    }
}

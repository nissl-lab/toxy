using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
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
			ClassicAssert.AreEqual("hello world"+Environment.NewLine+"a2"+Environment.NewLine+"a3"+Environment.NewLine+"bbb4"+Environment.NewLine, text);
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
                    ClassicAssert.AreEqual("hello world", args.Text);
                }
                else if(args.LineNumber==1)
                {
                    ClassicAssert.AreEqual("a2", args.Text);
                }
                else if (args.LineNumber == 2)
                {
                    ClassicAssert.AreEqual("a3", args.Text);
                }
                else if (args.LineNumber == 3)
                {
                    ClassicAssert.AreEqual("bbb4", args.Text);
                }
            };
            string text = parser.Parse();
            ClassicAssert.IsNotNull(text);
            ClassicAssert.IsNotEmpty(text);
        }

    }
}

using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

namespace Toxy.Test
{
    [TestFixture]
    public class Powerpoint2007SlideshowParserTest
    {
        [Test]
        public void ReadTextBasicTest()
        {
            string path = Path.GetFullPath(TestDataSample.GetPowerpointPath("testPPT.pptx"));
            ParserContext context = new ParserContext(path);
            ISlideshowParser parser = ParserFactory.CreateSlideshow(context);
            var result = parser.Parse();
            ClassicAssert.AreEqual(3, result.Slides.Count);
            ClassicAssert.AreEqual(2, result.Slides[0].Texts.Count);
            ClassicAssert.AreEqual("Attachment Test", result.Slides[0].Texts[0]);
            ClassicAssert.AreEqual("Rajiv", result.Slides[0].Texts[1]);
            ClassicAssert.AreEqual(6, result.Slides[1].Texts.Count);
            ClassicAssert.AreEqual("This is a test file data with the same content as every other file being tested for ",
                result.Slides[1].Texts[0]);
            ClassicAssert.AreEqual("tika", result.Slides[1].Texts[1]);
            ClassicAssert.AreEqual("Kumar ", result.Slides[1].Texts[3]);

            ClassicAssert.AreEqual(10, result.Slides[2].Texts.Count);
            ClassicAssert.AreEqual("Different words to test against", result.Slides[2].Texts[0]);
        }
        [Test]
        public void ReadTextBySlideIndex()
        {
            string path = Path.GetFullPath(TestDataSample.GetPowerpointPath("testPPT.pptx"));
            ParserContext context = new ParserContext(path);
            ISlideshowParser parser = ParserFactory.CreateSlideshow(context);
            var result = parser.Parse(1);
            ClassicAssert.AreEqual(6, result.Texts.Count);
            ClassicAssert.AreEqual("This is a test file data with the same content as every other file being tested for ",
                result.Texts[0]);
            ClassicAssert.AreEqual("tika", result.Texts[1]);
            ClassicAssert.AreEqual(" content parsing. This has been developed by Rajiv ", result.Texts[2]);    
            ClassicAssert.AreEqual("Kumar ", result.Texts[3]);            
        }
    }
}
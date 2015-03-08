using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
            Assert.AreEqual(3, result.Slides.Count);
            Assert.AreEqual(2, result.Slides[0].Texts.Count);
            Assert.AreEqual("Attachment Test", result.Slides[0].Texts[0]);
            Assert.AreEqual("Rajiv", result.Slides[0].Texts[1]);
            Assert.AreEqual(6, result.Slides[1].Texts.Count);
            Assert.AreEqual("This is a test file data with the same content as every other file being tested for ",
                result.Slides[1].Texts[0]);
            Assert.AreEqual("tika", result.Slides[1].Texts[1]);
            Assert.AreEqual("Kumar ", result.Slides[1].Texts[3]);

            Assert.AreEqual(10, result.Slides[2].Texts.Count);
            Assert.AreEqual("Different words to test against", result.Slides[2].Texts[0]);
        }
        [Test]
        public void ReadTextBySlideIndex()
        {
            string path = Path.GetFullPath(TestDataSample.GetPowerpointPath("testPPT.pptx"));
            ParserContext context = new ParserContext(path);
            ISlideshowParser parser = ParserFactory.CreateSlideshow(context);
            var result = parser.Parse(1);
            Assert.AreEqual(6, result.Texts.Count);
            Assert.AreEqual("This is a test file data with the same content as every other file being tested for ",
                result.Texts[0]);
            Assert.AreEqual("tika", result.Texts[1]);
            Assert.AreEqual(" content parsing. This has been developed by Rajiv ", result.Texts[2]);    
            Assert.AreEqual("Kumar ", result.Texts[3]);            
        }
    }
}
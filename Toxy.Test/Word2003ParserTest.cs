using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class Word2003ParserTest
    {

        [Test]
        public void TestParseTextFromWord()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.doc"));
            ITextParser parser = ParserFactory.CreateText(context);
            string doc = parser.Parse();

            Assert.IsNotNull(doc);

            string[] lines = doc.Split('\r');
            ClassicAssert.AreEqual(8, lines.Length);
            ClassicAssert.AreEqual("I am a test document", lines[0]);
            ClassicAssert.AreEqual("This is page 1", lines[1]);
            ClassicAssert.AreEqual("I am Calibri (Body) in font size 11", lines[2]);
            ClassicAssert.AreEqual("\f", lines[3]);
            ClassicAssert.AreEqual("This is page two", lines[4]);
            ClassicAssert.AreEqual("It’s Arial Black in 16 point", lines[5]);
            ClassicAssert.AreEqual("It’s also in blue", lines[6]);
            ClassicAssert.AreEqual("", lines[7]);
        }
        [Test]
        public void TestParseSimpleDocumentFromWord()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.doc"));
            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument doc = parser.Parse();
            ClassicAssert.AreEqual(8,doc.Paragraphs.Count);
            ClassicAssert.AreEqual("I am a test document\r",doc.Paragraphs[0].Text);
            ClassicAssert.AreEqual("This is page 1\r", doc.Paragraphs[1].Text);
            ClassicAssert.AreEqual("I am Calibri (Body) in font size 11\r", doc.Paragraphs[2].Text);
            ClassicAssert.AreEqual("\f", doc.Paragraphs[3].Text);
            ClassicAssert.AreEqual("\r", doc.Paragraphs[4].Text);
            ClassicAssert.AreEqual("This is page two\r", doc.Paragraphs[5].Text);
            ClassicAssert.AreEqual("It’s Arial Black in 16 point\r", doc.Paragraphs[6].Text);
            ClassicAssert.AreEqual("It’s also in blue\r", doc.Paragraphs[7].Text);
        }
    }
}

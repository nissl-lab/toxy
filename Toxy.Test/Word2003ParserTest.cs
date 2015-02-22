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
            Assert.AreEqual(8, lines.Length);
            Assert.AreEqual("I am a test document", lines[0]);
            Assert.AreEqual("This is page 1", lines[1]);
            Assert.AreEqual("I am Calibri (Body) in font size 11", lines[2]);
            Assert.AreEqual("\f", lines[3]);
            Assert.AreEqual("This is page two", lines[4]);
            Assert.AreEqual("It’s Arial Black in 16 point", lines[5]);
            Assert.AreEqual("It’s also in blue", lines[6]);
            Assert.AreEqual("", lines[7]);
        }
        [Test]
        public void TestParseSimpleDocumentFromWord()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.doc"));
            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument doc = parser.Parse();
            Assert.AreEqual(8,doc.Paragraphs.Count);
            Assert.AreEqual("I am a test document\r",doc.Paragraphs[0].Text);
            Assert.AreEqual("This is page 1\r", doc.Paragraphs[1].Text);
            Assert.AreEqual("I am Calibri (Body) in font size 11\r", doc.Paragraphs[2].Text);
            Assert.AreEqual("\f", doc.Paragraphs[3].Text);
            Assert.AreEqual("\r", doc.Paragraphs[4].Text);
            Assert.AreEqual("This is page two\r", doc.Paragraphs[5].Text);
            Assert.AreEqual("It’s Arial Black in 16 point\r", doc.Paragraphs[6].Text);
            Assert.AreEqual("It’s also in blue\r", doc.Paragraphs[7].Text);
        }
    }
}

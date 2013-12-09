using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestClass]
    public class Word2007ParserTest
    {
        [TestMethod]
        public void TestParseSimpleDocument()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.docx"));
            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument doc = parser.Parse();
            Assert.AreEqual(7,doc.Paragraphs.Count);
            Assert.AreEqual("I am a test document",doc.Paragraphs[0].Text);
            Assert.AreEqual("This is page 1", doc.Paragraphs[1].Text);
            Assert.AreEqual("I am Calibri (Body) in font size 11", doc.Paragraphs[2].Text);
            Assert.AreEqual("\n", doc.Paragraphs[3].Text);
            Assert.AreEqual("This is page two", doc.Paragraphs[4].Text);
            Assert.AreEqual("It’s Arial Black in 16 point", doc.Paragraphs[5].Text);
            Assert.AreEqual("It’s also in blue", doc.Paragraphs[6].Text);
        }
    }
}

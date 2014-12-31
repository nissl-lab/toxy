using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Toxy.Parsers;

namespace Toxy.Test
{
    [TestFixture]
    public class PDFParserTest
    {
        [Test]
        public void TestParsePlainTextFromPDF()
        {
            string path = TestDataSample.GetPdfPath("Sample1.PDF");
            var parser = new PDFTextParser(new ParserContext(path));
            string result = parser.Parse();
            Assert.IsTrue(result.StartsWith("LA MARCHE"));
        }

        [Test]
        public void TestParseToxyDocumentFromPDF()
        {
            string path = TestDataSample.GetPdfPath("Sample1.PDF");
            var parser = new PDFDocumentParser(new ParserContext(path));
            var result = parser.Parse();
            Assert.AreEqual(1474, result.Paragraphs.Count);
            Assert.AreEqual("LA MARCHE MONDIALE DES FEMMES : UN MOUVEMENT IRRÉVERSIBLE", result.Paragraphs[0].Text);
        }
    }
}

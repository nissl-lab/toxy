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
            Assert.AreEqual(88, result.Paragraphs.Count);
            string[] results=result.Paragraphs[0].Text.Split('\n');
            Assert.AreEqual("LA MARCHE MONDIALE DES FEMMES : UN MOUVEMENT IRRÉVERSIBLE", results[0]);
            Assert.AreEqual("DOCUMENT PRÉPARATOIRE", results[1]);
            Assert.AreEqual("4eRencontre internationale de la Marche mondiale des femmes", results[2]);
            Assert.AreEqual("du 18-22 Mars 2003", results[3]);
        }
    }
}

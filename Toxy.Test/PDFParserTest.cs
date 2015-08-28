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
        void ContainText(string result, string text)
        {
            Assert.IsTrue(result.IndexOf(text) > 0);
        }
        [Test]
        public void TestParsePlainTextFromSample1()
        {
            string path = TestDataSample.GetPdfPath("Sample1.PDF");
            var parser = new PDFTextParser(new ParserContext(path));
            string result = parser.Parse();
            Assert.IsTrue(result.StartsWith("LA MARCHE"));
            ContainText(result, "Toute discussion stratégique sur nos actions nécessite un rappel de ce que nous avons fait en");
            ContainText(result, "l’an 2000 et depuis. Au niveau mondial, en l’an 2000, nous avons mené une campagne de");
            ContainText(result, "Une structure pour nous amener à 2005");
            ContainText(result, "Lors de la 4e rencontre qui aura lieu en Inde, nous avons deux objectifs majeurs");
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
        [Test]
        public void TestParsePlainTextFromPDF2()
        {
            string path = TestDataSample.GetPdfPath("Sample5.PDF");
            var parser = new PDFTextParser(new ParserContext(path));
            string result = parser.Parse();
            Assert.IsTrue(result.StartsWith("Philadelphia, Atlanta, Dallas, San Diego, and New Orleans."));


        }
    }
}

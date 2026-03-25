using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using Toxy;
using Toxy.Parsers;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Exceptions;

namespace Toxy.Test
{
    [TestFixture]
    public class PDFParserTest
    {
        void ContainText(string result, string text)
        {
            ClassicAssert.IsTrue(result.IndexOf(text) > 0);
        }
        [Test]
        public void TestParsePlainTextFromSample1()
        {
            string path = TestDataSample.GetPdfPath("Sample1.PDF");
            var parser = new PDFTextParser(new ParserContext(path));
            string result = parser.Parse();
            ClassicAssert.IsTrue(result.StartsWith("LA MARCHE"));
            ContainText(result, "Toute discussion stratégique sur nos actions nécessite un rappel de ce que nous avons fait en");
            ContainText(result, "l'an 2000 et depuis. Au niveau mondial, en l'an 2000, nous avons mené une campagne de");
            ContainText(result, "Une structure pour nous amener à 2005");
            ContainText(result, "Lors de la 4e rencontre qui aura lieu en Inde, nous avons deux objectifs majeurs");
        }

        [Test]
        public void TestParseToxyDocumentFromPDF()
        {
            string path = TestDataSample.GetPdfPath("Sample1.PDF");
            var parser = new PDFDocumentParser(new ParserContext(path));
            var result = parser.Parse();
            //ClassicAssert.AreEqual(286, result.Paragraphs.Count);
            ClassicAssert.IsTrue(result.Paragraphs[0].Text.StartsWith("LAMARCHEMONDIALE DES FEMMES : UN MOUVEMENT IRRÉVERSIBLE"));
        }
        [Test]
        public void TestParsePlainTextFromSample5()
        {
            string path = TestDataSample.GetPdfPath("Sample5.PDF");
            var parser = new PDFTextParser(new ParserContext(path));
            string result = parser.Parse();
            string[] lines = result.Split(Environment.NewLine, System.StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual("Philadelphia, Atlanta, Dallas, San Diego, and New", lines[0]);
            ClassicAssert.AreEqual("Orleans. According tocompanyestimates, its own sales", lines[1]);
        }
        [Test]
        public void TestReadBigPDFFile()
        {
            string path = TestDataSample.GetPdfPath("Word97-2007BinaryFileFormat(doc)Specification.pdf");
            var parser = new PDFTextParser(new ParserContext(path));
            ClassicAssert.DoesNotThrow(() => parser.Parse());
        }
        [Test]
        public void TestStreamForPDFTextParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("Sample1.PDF", "PDF"));
            var parser = ParserFactory.CreateText(context);
            string result = parser.Parse();
        }
        [Test]
        public void TestStreamForPDFDocumentParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("Sample1.PDF", "PDF"));
            var parser = ParserFactory.CreateDocument(context);
            var result = parser.Parse();
        }
        [Test]
        public void TestEncryptedPDF()
        {
            string path = TestDataSample.GetPdfPath("password.pdf");
            var parser = new PDFTextParser(new ParserContext(path));
            ClassicAssert.Throws<PdfDocumentEncryptedException>(() => parser.Parse());
        }

        [Test]
        public void TestParseSpreadsheetFromPDF()
        {
            string path = TestDataSample.GetPdfPath("xlsdemo1.pdf");
            var parser = new PDFSpreadsheetParser(new ParserContext(path));
            var result = parser.Parse();
            ClassicAssert.IsNotNull(result);
            ClassicAssert.IsTrue(result.Tables.Count > 0);
        }

        [Test]
        public void TestParseSpreadsheetFromPDFStream()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("xlsdemo1.pdf", "PDF"));
            var parser = ParserFactory.CreateSpreadsheet(context);
            var result = parser.Parse();
            ClassicAssert.IsNotNull(result);
        }

        [Test]
        public void TestParseSpreadsheetFromPDFSecondFile()
        {
            string path = TestDataSample.GetPdfPath("xlsdemo2.pdf");
            var parser = new PDFSpreadsheetParser(new ParserContext(path));
            var result = parser.Parse();
            ClassicAssert.IsNotNull(result);
        }
    }
}

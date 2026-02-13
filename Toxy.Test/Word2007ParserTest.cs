using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;

namespace Toxy.Test
{
    [TestFixture]
    public class Word2007ParserTest
    {

        [Test]
        public void TestParseTextFromWord()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.docx"));
            ITextParser parser = ParserFactory.CreateText(context);
            string doc = parser.Parse();

            ClassicAssert.IsNotNull(doc);

            string[] lines = doc.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual(7, lines.Length);
            ClassicAssert.AreEqual("I am a test document", lines[0]);
            ClassicAssert.AreEqual("This is page 1", lines[1]);
            ClassicAssert.AreEqual("I am Calibri (Body) in font size 11", lines[2]);
            ClassicAssert.AreEqual("This is page two", lines[4]);
            ClassicAssert.AreEqual("It’s Arial Black in 16 point", lines[5]);
            ClassicAssert.AreEqual("It’s also in blue", lines[6]);
        }
        [Test]
        public void TestParseSimpleDocumentFromWord()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("SampleDoc.docx"));
            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument doc = parser.Parse();
            ClassicAssert.AreEqual(7,doc.Paragraphs.Count);
            ClassicAssert.AreEqual("I am a test document",doc.Paragraphs[0].Text);
            ClassicAssert.AreEqual("This is page 1", doc.Paragraphs[1].Text);
            ClassicAssert.AreEqual("I am Calibri (Body) in font size 11", doc.Paragraphs[2].Text);
            ClassicAssert.AreEqual("\n", doc.Paragraphs[3].Text);
            ClassicAssert.AreEqual("This is page two", doc.Paragraphs[4].Text);
            ClassicAssert.AreEqual("It’s Arial Black in 16 point", doc.Paragraphs[5].Text);
            ClassicAssert.AreEqual("It’s also in blue", doc.Paragraphs[6].Text);
        }
        [Test]
        public void TestParseDocumentWithTable()
        {
            ParserContext context = new ParserContext(TestDataSample.GetWordPath("simple-table.docx"));
            IDocumentParser parser = ParserFactory.CreateDocument(context);
            ToxyDocument doc = parser.Parse();
            ClassicAssert.AreEqual(8, doc.Paragraphs.Count);
            ClassicAssert.AreEqual("This is a Word document that was created using Word 97 – SR2.  It contains a paragraph, a table consisting of 2 rows and 3 columns and a final paragraph.", 
                doc.Paragraphs[0].Text);
            ClassicAssert.AreEqual("This text is below the table.", doc.Paragraphs[1].Text);
            ClassicAssert.AreEqual("Cell 1,1", doc.Paragraphs[2].Text);
            ClassicAssert.AreEqual("Cell 1,2", doc.Paragraphs[3].Text);
            ClassicAssert.AreEqual("Cell 1,3", doc.Paragraphs[4].Text);
            ClassicAssert.AreEqual("Cell 2,1", doc.Paragraphs[5].Text);
            ClassicAssert.AreEqual("Cell 2,2", doc.Paragraphs[6].Text);
            ClassicAssert.AreEqual("Cell 2,3", doc.Paragraphs[7].Text);
        }
    }
}

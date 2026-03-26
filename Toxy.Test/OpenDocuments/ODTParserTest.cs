using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

namespace Toxy.Test.OpenDocuments
{
	public class ODTParserTest
	{
		[Test]
		public void TestParseTextFromODP()
		{
			ParserContext context = new ParserContext(TestDataSample.GetODTPath("SampleDoc.odt"));
			ITextParser parser = ParserFactory.CreateText(context);
			string doc = parser.Parse();

			ClassicAssert.IsNotNull(doc);

			ClassicAssert.IsTrue(doc.Contains("I am a test document"));
			ClassicAssert.IsTrue(doc.Contains("This is page 1"));
			ClassicAssert.IsTrue(doc.Contains("I am Calibri (Body) in font size 11"));
			ClassicAssert.IsTrue(doc.Contains("This is page two"));
			ClassicAssert.IsTrue(doc.Contains("It’s Arial Black in 16 point"));
			ClassicAssert.IsTrue(doc.Contains("It’s also in blue"));
		}
		[Test]
		public void TestStreamForODPTextParser()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODTPath("SampleDoc.odt"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				ITextParser parser = ParserFactory.CreateText(context);
				string result = parser.Parse();
			}
		}
	}
}

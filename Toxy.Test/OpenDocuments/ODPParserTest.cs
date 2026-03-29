using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Toxy.Test.OpenDocuments
{
	[TestFixture]
	public class ODPParserTest
	{
		[Test]
		public void ReadTextBasicTest()
		{
			string path = Path.GetFullPath(TestDataSample.GetODPPath("testPPT.odp"));
			ParserContext context = new ParserContext(path);
			ITextParser parser = ParserFactory.CreateText(context);
			string result = parser.Parse();
			ClassicAssert.IsNotNull(result);
			ClassicAssert.IsNotEmpty(result);
			string[] texts = result.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
			ClassicAssert.AreEqual(14, texts.Length);
			ClassicAssert.AreEqual("Attachment Test", texts[0]);
			ClassicAssert.AreEqual("Rajiv", texts[1]);
			ClassicAssert.AreEqual("Different words to test against", texts[4]);
			ClassicAssert.AreEqual("Hello", texts[7]);
		}
		[Test]
		public void TestStreamForTextParser()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODPPath("testPPT.odp"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				ITextParser parser = ParserFactory.CreateText(context);
				string result = parser.Parse();
			}
		}

		[Test]
		public void TestMetadata()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODPPath("SampleShow.odp"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				IMetadataParser parser = ParserFactory.CreateMetadata(context);
				ToxyMetadata result = parser.Parse();
				ClassicAssert.AreEqual("word", result.Get("Creator").Value);
				ClassicAssert.AreEqual(DateTimeOffset.Parse("2026-03-26T19:45:15"), result.Get("Date").Value);
				ClassicAssert.AreEqual("This is a sample slideshow, for use with testing etc", result.Get("Description").Value);
				ClassicAssert.AreEqual("", result.Get("Language").Value);
				ClassicAssert.AreEqual("A sample slideshow", result.Get("Subject").Value);
				ClassicAssert.AreEqual("SlideShow Sample", result.Get("Title").Value);
				ClassicAssert.AreEqual("", result.Get("Contributor").Value);
				ClassicAssert.AreEqual("", result.Get("Coverage").Value);
				ClassicAssert.AreEqual("", result.Get("Identifier").Value);
				ClassicAssert.AreEqual("", result.Get("Publisher").Value);
				ClassicAssert.AreEqual("", result.Get("Relation").Value);
				ClassicAssert.AreEqual("", result.Get("Rights").Value);
				ClassicAssert.AreEqual("", result.Get("Source").Value);
				ClassicAssert.AreEqual("", result.Get("Type").Value);

				ClassicAssert.AreEqual(DateTimeOffset.Parse("2008-01-04T11:58:07Z"), result.Get("CreationDate").Value);
				ClassicAssert.AreEqual(4, result.Get("EditingCycles").Value);
				ClassicAssert.AreEqual("PT120S", result.Get("EditingDuration").Value);
				ClassicAssert.AreEqual("MicrosoftOffice/14.0 MicrosoftPowerPoint", result.Get("Generator").Value);
				ClassicAssert.AreEqual("Nick Burch", result.Get("InitialCreator").Value);
				ClassicAssert.AreEqual("", result.Get("PrintedBy").Value);
				ClassicAssert.AreEqual(DateTimeOffset.MinValue, result.Get("PrintDate").Value);
				ClassicAssert.AreEqual("application/vnd.oasis.opendocument.presentation", result.Get("MimeType").Value);

				List<string> keywords = result.Get("Keywords").Value as List<string>;
				ClassicAssert.AreEqual("Sample Testing", keywords.First());
			}
		}
	}
}

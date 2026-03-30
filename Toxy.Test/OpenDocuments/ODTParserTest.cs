using DocumentFormat.OpenXml.Office2010.ExcelAc;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Toxy.Test.OpenDocuments
{
	[TestFixture]
	public class ODTParserTest
	{
		[Test]
		public void TestParseTextFromODT()
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
		public void TestStreamForODTTextParser()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODTPath("SampleDoc.odt"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				ITextParser parser = ParserFactory.CreateText(context);
				string result = parser.Parse();
			}
		}

		[Test]
		public void TestMetadata()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODTPath("F2013C00907VOL04.odt"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				IMetadataParser parser = ParserFactory.CreateMetadata(context);
				ToxyMetadata result = parser.Parse();
				ClassicAssert.AreEqual("", result.Get("Creator").Value);
				ClassicAssert.AreEqual(DateTimeOffset.Parse("2026-03-28T08:14:06.3500992"), result.Get("Date").Value);
				ClassicAssert.AreEqual("This is my comment", result.Get("Description").Value);
				ClassicAssert.AreEqual("", result.Get("Language").Value);
				ClassicAssert.AreEqual("My Topic", result.Get("Subject").Value);
				ClassicAssert.AreEqual("Schedules 2 and 3", result.Get("Title").Value);
				ClassicAssert.AreEqual("Additional Authors", result.Get("Contributor").Value);
				ClassicAssert.AreEqual("This is my summary", result.Get("Coverage").Value);
				ClassicAssert.AreEqual("This is my Identifier", result.Get("Identifier").Value);
				ClassicAssert.AreEqual("I did not published it", result.Get("Publisher").Value);
				ClassicAssert.AreEqual("my relation", result.Get("Relation").Value);
				ClassicAssert.AreEqual("my rights", result.Get("Rights").Value);
				ClassicAssert.AreEqual("my source", result.Get("Source").Value);
				ClassicAssert.AreEqual("My type", result.Get("Type").Value);

				ClassicAssert.AreEqual(DateTimeOffset.Parse("2011-02-10T02:16:00Z"), result.Get("CreationDate").Value);
				ClassicAssert.AreEqual(2, result.Get("EditingCycles").Value);
				ClassicAssert.AreEqual("PT13H37M2S", result.Get("EditingDuration").Value);
				ClassicAssert.AreEqual("LibreOffice/26.2.0.3$Windows_X86_64 LibreOffice_project/620$Build-3", result.Get("Generator").Value);
				ClassicAssert.AreEqual("urquhm", result.Get("InitialCreator").Value);
				ClassicAssert.AreEqual("", result.Get("PrintedBy").Value);
				ClassicAssert.AreEqual(DateTimeOffset.Parse("2013-09-19T00:38:00Z"), result.Get("PrintDate").Value);
				ClassicAssert.AreEqual("application/vnd.oasis.opendocument.text", result.Get("MimeType").Value);

				List<string> keywords = result.Get("Keywords").Value as List<string>;
				ClassicAssert.AreEqual("My Keywords", keywords.First());
			}
		}
	}
}

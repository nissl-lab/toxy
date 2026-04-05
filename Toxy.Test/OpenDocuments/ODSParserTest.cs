using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Toxy.Parsers.OpenDocument.Entities;

namespace Toxy.Test.OpenDocuments
{
	[TestFixture]
	public class ODSParserTest
	{
		[Test]
		public void TestODSTextParser()
		{
			ParserContext context = new ParserContext(TestDataSample.GetODSPath("Employee.ods"));
			ITextParser parser = ParserFactory.CreateText(context);
			string result = parser.Parse();
			ClassicAssert.IsNotNull(result);
			ClassicAssert.IsTrue(result.IndexOf("Last name") > 0);
			ClassicAssert.IsTrue(result.IndexOf("First name") > 0);
		}
		[Test]
		public void TestODSTextParser2()
		{
			ParserContext context = new ParserContext(TestDataSample.GetODSPath("WithVariousData.ods"));
			ITextParser parser = ParserFactory.CreateText(context);
			string result = parser.Parse();
			ClassicAssert.IsNotNull(result);
			ClassicAssert.IsTrue(result.IndexOf("Foo") > 0);
			ClassicAssert.IsTrue(result.IndexOf("Bar") > 0);
			ClassicAssert.IsTrue(result.IndexOf("a really long cell") > 0);
			ClassicAssert.IsTrue(result.IndexOf("have a header") > 0);
			ClassicAssert.IsTrue(result.IndexOf("have a footer") > 0);
			ClassicAssert.IsTrue(result.IndexOf("This is the header") < 0);
		}
		[Test]
		public void TestODSTextParserWithoutSheetNames()
		{
			ParserContext context = new ParserContext(TestDataSample.GetODSPath("WithVariousData.ods"));
			context.Properties.Add("IncludeSheetNames", "0");
			ITextParser parser = ParserFactory.CreateText(context);
			string result = parser.Parse();
			ClassicAssert.IsNotNull(result);
			ClassicAssert.IsTrue(result.IndexOf("Sheet1") < 0);
		}

		[Test]
		public void TestStreamForTextParser()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODSPath("WithVariousData.ods"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				ITextParser parser = ParserFactory.CreateText(context);
				string result = parser.Parse();
			}
		}

		[Test]
		public void TestMetadata()
		{
			using (FileStream fs = new FileStream(TestDataSample.GetODSPath("WithVariousData.ods"), FileMode.Open))
			{
				ParserContext context = new ParserContext(fs);
				IMetadataParser parser = ParserFactory.CreateMetadata(context);
				ToxyMetadata result = parser.Parse();
				ClassicAssert.AreEqual("", result.Get("Creator").Value);
				ClassicAssert.AreEqual(DateTimeOffset.Parse("2026-03-28T09:20:23.7074766"), result.Get("Date").Value);
				ClassicAssert.AreEqual("This is my comment", result.Get("Description").Value);
				ClassicAssert.AreEqual("", result.Get("Language").Value);
				ClassicAssert.AreEqual("My Topic", result.Get("Subject").Value);
				ClassicAssert.AreEqual("My Title", result.Get("Title").Value);
				ClassicAssert.AreEqual("Additional Authors", result.Get("Contributor").Value);
				ClassicAssert.AreEqual("This is my summary", result.Get("Coverage").Value);
				ClassicAssert.AreEqual("This is my Identifier", result.Get("Identifier").Value);
				ClassicAssert.AreEqual("I did not published it", result.Get("Publisher").Value);
				ClassicAssert.AreEqual("my relation", result.Get("Relation").Value);
				ClassicAssert.AreEqual("my rights", result.Get("Rights").Value);
				ClassicAssert.AreEqual("my source", result.Get("Source").Value);
				ClassicAssert.AreEqual("My type", result.Get("Type").Value);

				ClassicAssert.AreEqual(DateTimeOffset.Parse("2008-03-31T21:49:23Z"), result.Get("CreationDate").Value);
				ClassicAssert.AreEqual(2, result.Get("EditingCycles").Value);
				ClassicAssert.AreEqual("PT4M42S", result.Get("EditingDuration").Value);
				ClassicAssert.AreEqual("LibreOffice/26.2.0.3$Windows_X86_64 LibreOffice_project/620$Build-3", result.Get("Generator").Value);
				ClassicAssert.AreEqual("Nick Burch", result.Get("InitialCreator").Value);
				ClassicAssert.AreEqual("", result.Get("PrintedBy").Value);
				ClassicAssert.AreEqual(DateTimeOffset.MinValue, result.Get("PrintDate").Value);
				ClassicAssert.AreEqual("application/vnd.oasis.opendocument.spreadsheet", result.Get("MimeType").Value);
				
				List<string> keywords = result.Get("Keywords").Value as List<string>;
				ClassicAssert.AreEqual("My Keywords", keywords.First());

				List<UserDefinedProperty> definedProperties = result.Get("UserDefinedProperties").Value as List<UserDefinedProperty>;
				ClassicAssert.AreEqual("true", definedProperties[0].Value);
				ClassicAssert.AreEqual("2026-03-28", definedProperties[1].Value);
				ClassicAssert.AreEqual("2026-03-28T00:00:19", definedProperties[2].Value);
				ClassicAssert.AreEqual("P36Y17M2DT6H8M4.000000021S", definedProperties[3].Value);
				ClassicAssert.AreEqual("4436346346", definedProperties[4].Value);
				ClassicAssert.AreEqual("My Property Text", definedProperties[5].Value);
			}
		}
	}
}

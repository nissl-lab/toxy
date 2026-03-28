using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

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
	}
}

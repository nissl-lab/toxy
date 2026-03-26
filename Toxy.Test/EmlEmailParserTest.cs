using NUnit.Framework;
using NUnit.Framework.Legacy;
using System.IO;

namespace Toxy.Test
{
    [TestFixture]
    public class EmlEmailParserTest
    {
		void ContainText(string result, string text)
		{
			ClassicAssert.IsTrue(result.IndexOf(text) > 0, result);
		}

		[Test]
        public void ReadEmlTest()
        {
            string path = TestDataSample.GetEmailPath("test.eml");
            ParserContext context = new ParserContext(path);
            IEmailParser parser = ParserFactory.CreateEmail(context);
            ToxyEmail email = parser.Parse();
            ClassicAssert.IsNotNull(email.From);
            ClassicAssert.IsNotEmpty(email.From);
            ClassicAssert.AreEqual(1, email.To.Count);
            ClassicAssert.AreEqual("\"拉勾网\" <service@email.lagou.com>", email.From);
            ClassicAssert.AreEqual("tonyqus@163.com", email.To[0]);
            ClassicAssert.IsTrue(email.Subject.StartsWith("上海分众德峰广告传播有限公司-高级.NET软件工程师招聘反馈通知"));
            ClassicAssert.IsNotNull(email.HtmlBody);
            ClassicAssert.IsNull(email.TextBody);
        }

		[Test]
		public void ReadEmlTextTest()
		{
			string path = TestDataSample.GetEmailPath("test.eml");
			ParserContext context = new ParserContext(path);
			ITextParser parser = ParserFactory.CreateText(context);
			string text = parser.Parse();
			ClassicAssert.IsNotNull(text);
            ContainText(text, "\"拉勾网\" <service@email.lagou.com>");
            ContainText(text, "tonyqus@163.com");
            ContainText(text, "上海分众德峰广告传播有限公司-高级.NET软件工程师招聘反馈通知");
		}

		[Test]
		public void ReadEmlMetaTest()
		{
			string path = TestDataSample.GetEmailPath("test.eml");
			ParserContext context = new ParserContext(path);
            IMetadataParser parser = ParserFactory.CreateMetadata(context);
			ToxyMetadata meta = parser.Parse();
			ClassicAssert.IsNotNull(meta);
			ClassicAssert.AreEqual("\"拉勾网\" <service@email.lagou.com>", meta.Get("From").Value.ToString());
			ClassicAssert.AreEqual("上海分众德峰广告传播有限公司-高级.NET软件工程师招聘反馈通知", meta.Get("Subject").Value.ToString());
			ClassicAssert.AreEqual("tonyqus@163.com", meta.Get("To").Value.ToString());
		}


		[Test]
        public void TestStreamForEmlTextParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("test.eml", "Email"));
            Assert.Throws<InvalidDataException>(() => { var parser = ParserFactory.CreateText(context); });
        }
        [Test]
        public void TestStreamForEmlEmailParser()
        {
            ParserContext context = new ParserContext(TestDataSample.GetFileStream("test.eml", "Email"));
            Assert.Throws<InvalidDataException>(() => { var parser = ParserFactory.CreateEmail(context); });
        }
    }
}

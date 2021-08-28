using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class MsgEmailParserTest
    {
        [Test]
        public void HtmlMsg_ReadTextTest()
        {
            string path = TestDataSample.GetEmailPath("Azure pricing and services updates.msg");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateText(context);

            string result=parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsNotEmpty(result);
            Assert.IsTrue(result.IndexOf("[From] Azure Team<AzureTeam@e-mail.microsoft.com>") >= 0);
            Assert.IsTrue(result.IndexOf("[To] tonyqux@hotmail.com")>0);
            Assert.IsTrue(result.IndexOf("[Subject] Azure pricing and services updates") > 0);
            Assert.IsFalse(result.IndexOf("[Cc]") > 0);
            Assert.IsFalse(result.IndexOf("[Bcc]") > 0);
        }

        [Test]
        public void PureTextMsg_ReadTextTest()
        {
            string path = TestDataSample.GetEmailPath("raw text mail demo.msg");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateText(context);

            string result = parser.Parse();
            Assert.IsNotNull(result);
            Assert.IsNotEmpty(result);
        }

        [Test]
        public void HtmlMsg_ReadMsgEntityTest()
        {
            //patch for 'No data is available for encoding 936'
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            string path = TestDataSample.GetEmailPath("Azure pricing and services updates.msg");
            ParserContext context = new ParserContext(path);
            var parser = ParserFactory.CreateEmail(context);

            var result = parser.Parse();
            Assert.AreEqual("Azure Team<AzureTeam@e-mail.microsoft.com>",result.From);

            Assert.AreEqual(1, result.To.Count);
            Assert.AreEqual(0, result.Cc.Count);
            Assert.AreEqual(0, result.Bcc.Count);

            Assert.AreEqual("tonyqux@hotmail.com", result.To[0]);
            Assert.AreEqual("Azure pricing and services updates", result.Subject);
            Assert.IsNotNull(result.TextBody);
            Assert.IsNotNull(result.HtmlBody);
        }
    }
}

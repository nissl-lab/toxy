using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class CnmEmailParserTest
    {

        [Test]
        public void ReadCnmTest()
        {
            string path = TestDataSample.GetCNMPath("Y0B6E8H2.CNM");
            ParserContext context = new ParserContext(path);
            IEmailParser parser = ParserFactory.CreateEmail(context) as IEmailParser;
            ToxyEmail email = parser.Parse();
            Assert.IsNotNullOrEmpty(email.From);
            Assert.AreEqual(1, email.To.Count);
            Assert.AreEqual("Тюльпаны <info@beepy.net>", email.From);
            Assert.AreEqual("<maf@1gb.ru>", email.To[0]);

            Assert.AreEqual("Тюльпаны", email.Subject);
            Assert.IsTrue(email.TextBody.StartsWith("Тел: 960-51-57;Продажа тюльпанов"));
            Assert.IsTrue(email.HtmlBody.StartsWith("<!DOCTYPE HTML"));

        }
    }
}

using System;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;

namespace Toxy.Test
{
    [TestFixture]
    public class HtmlParserTest
    {
        [Test]
        public void TestParseHtml()
        {
            string path = Path.GetFullPath(TestDataSample.GetHtmlPath("mshome.html"));

            ParserContext context = new ParserContext(path);
            IDomParser parser = (IDomParser)ParserFactory.CreateDom(context);
            ToxyDom toxyDom = parser.Parse();

            List<ToxyNode> metaNodeList = toxyDom.Root.SelectNodes("//meta");
            Assert.AreEqual(7, metaNodeList.Count);

            ToxyNode aNode = toxyDom.Root.SingleSelect("//a");
            Assert.AreEqual(1, aNode.Attributes.Count);
            Assert.AreEqual("href", aNode.Attributes[0].Name);
            Assert.AreEqual("http://www.microsoft.com/en/us/default.aspx?redir=true", aNode.Attributes[0].Value);

            ToxyNode titleNode = toxyDom.Root.ChildrenNodes[0].ChildrenNodes[0].ChildrenNodes[0];
            Assert.AreEqual("title", titleNode.Name);
            Assert.AreEqual("Microsoft Corporation", titleNode.ChildrenNodes[0].InnerText);

            ToxyNode metaNode = toxyDom.Root.ChildrenNodes[0].ChildrenNodes[0].ChildrenNodes[7];
            Assert.AreEqual("meta", metaNode.Name);
            Assert.AreEqual(3, metaNode.Attributes.Count);
            Assert.AreEqual("name", metaNode.Attributes[0].Name);
            Assert.AreEqual("SearchDescription", metaNode.Attributes[0].Value);
            Assert.AreEqual("scheme", metaNode.Attributes[2].Name);
            Assert.AreEqual(string.Empty, metaNode.Attributes[2].Value);
        }
    }
}

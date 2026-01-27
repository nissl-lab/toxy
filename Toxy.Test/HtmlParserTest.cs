using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
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
            ClassicAssert.AreEqual(7, metaNodeList.Count);

            ToxyNode aNode = toxyDom.Root.SingleSelect("//a");
            ClassicAssert.AreEqual(1, aNode.Attributes.Count);
            ClassicAssert.AreEqual("href", aNode.Attributes[0].Name);
            ClassicAssert.AreEqual("http://www.microsoft.com/en/us/default.aspx?redir=true", aNode.Attributes[0].Value);

            ToxyNode titleNode = toxyDom.Root.ChildrenNodes[0].ChildrenNodes[0].ChildrenNodes[0];
            ClassicAssert.AreEqual("title", titleNode.Name);
            ClassicAssert.AreEqual("Microsoft Corporation", titleNode.ChildrenNodes[0].InnerText);

            ToxyNode metaNode = toxyDom.Root.ChildrenNodes[0].ChildrenNodes[0].ChildrenNodes[7];
            ClassicAssert.AreEqual("meta", metaNode.Name);
            ClassicAssert.AreEqual(3, metaNode.Attributes.Count);
            ClassicAssert.AreEqual("name", metaNode.Attributes[0].Name);
            ClassicAssert.AreEqual("SearchDescription", metaNode.Attributes[0].Value);
            ClassicAssert.AreEqual("scheme", metaNode.Attributes[2].Name);
            ClassicAssert.AreEqual(string.Empty, metaNode.Attributes[2].Value);
        }
    }
}

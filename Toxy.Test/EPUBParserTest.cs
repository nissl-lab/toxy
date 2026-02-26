using DocumentFormat.OpenXml.Office2010.ExcelAc;
using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using Toxy.Parsers;
using Toxy.Parsers.EPUB;

namespace Toxy.Test
{
    [TestFixture]
    public class EPUBParserTest
    {
        void ContainText(string result, string text)
        {
            ClassicAssert.IsTrue(result.IndexOf(text) > 0, result);
        }
        [Test]
        public void TestParseTextSample1()
        {
            string path = TestDataSample.GetEPUBPath("Sample1.epub");
            var parser = new EPUBTextParser(new ParserContext(path));
            string result = parser.Parse();
            string[] lines = result.Split('\n');
            ClassicAssert.Greater(lines.Length, 1);
            ClassicAssert.IsTrue(lines[1].StartsWith("LA MARCHE"));
            ContainText(result, "Toute discussion stratégique sur nos actions nécessite un rappel de ce que nous avons fait en");
            ContainText(result, "l’an 2000 et depuis. Au niveau mondial, en l’an 2000, nous avons mené une campagne de");
            ContainText(result, "Une structure pour nous amener à 2005");
            ContainText(result, "Lors de la 4e rencontre qui aura lieu en Inde, nous avons deux objectifs majeurs");
        }
        [Test]
        public void TestParseTextSample5()
        {
            string path = TestDataSample.GetEPUBPath("Sample5.epub");
            var parser = new EPUBTextParser(new ParserContext(path));
            string result = parser.Parse();
            string[] lines = result.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            ClassicAssert.AreEqual("License income by market (%) Philadelphia, Atlanta, Dallas, San Diego, and New", lines[1]);
            ClassicAssert.AreEqual("Orleans. According to company estimates, its own sales", lines[2]);
        }

        [Test]
        public void TestUnicodeTextContacts()
        {
            string path = TestDataSample.GetEPUBPath("contacts.epub");
            var parser = new EPUBTextParser(new ParserContext(path));
            string result = parser.Parse();
            ContainText(result, "云南国际信托有限公司内部联系方式");
            ContainText(result, "huyx@yntrust.com 南宁");
            ContainText(result, "（北京：会议电话：010-209");
            ContainText(result, "部门总经理助理");
            ContainText(result, "ranjz@yntrust.com 168");
            ContainText(result, "王艳芳");
            ContainText(result, "18616204608");
            ContainText(result, "wangq@yntrust.com 39");
            ContainText(result, "高级信息技术经理");
        }

        [Test]
        public void TestParseBigFile()
        {
            string path = TestDataSample.GetEPUBPath("Word97-2007BinaryFileFormat_doc_Specification.epub");
            var parser = new EPUBTextParser(new ParserContext(path));
            ClassicAssert.DoesNotThrow(() => parser.Parse());
        }

        [Test]
        public void TestMetaDataSample1()
        {
            string path = TestDataSample.GetEPUBPath("Metadata sample1.epub");
            var parser = new EPUBMetaParser(new ParserContext(path));
            var result = parser.Parse();

            ClassicAssert.AreEqual("高级信息技术经理, 高级信息技术经理", result.Get("Author")?.Value);

            List<string> authorList = new List<string>(2);
            if (result.Get("AuthorList")?.Value is List<string> temp)
            {
                authorList = temp;
            }
            ClassicAssert.AreEqual("高级信息技术经理", authorList[0]);
            ClassicAssert.AreEqual("高级信息技术经理", authorList[1]);
            ClassicAssert.AreEqual("高级信息技术经理", result.Get("Description")?.Value);
            ClassicAssert.AreEqual("高级信息技术经理", result.Get("Title")?.Value);
            
        }

        [Test]
        public void TestMetaDataSample2()
        {
            string path = TestDataSample.GetEPUBPath("Metadata sample2.epub");
            var parser = new EPUBMetaParser(new ParserContext(path));
            var result = parser.Parse();

            ClassicAssert.AreEqual("Unknown, Unknown", result.Get("Author")?.Value);

            List<string> authorList = new List<string>(2);
            if (result.Get("AuthorList")?.Value is List<string> temp)
            {
                authorList = temp;
            }
            ClassicAssert.AreEqual("Unknown", authorList[0]);
            ClassicAssert.AreEqual("Unknown", authorList[1]);
            ClassicAssert.AreEqual("Unknown", result.Get("Description")?.Value);
            ClassicAssert.AreEqual("Unknown", result.Get("Title")?.Value);
        }
    }
}

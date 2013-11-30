using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
     [TestClass]
    public class CsvParserTest
    {
         [TestMethod]
         public void TestParseToxySpreadsheet()
         {
             string path = TestDataSample.GetFilePath("countrylist.csv");

            ParserContext context=new ParserContext(path);
            context.Properties.Add("HasHeader", "1");
            ISpreadsheetParser parser = (ISpreadsheetParser)ParserFactory.Create(path);
            ToxySpreadsheet ss= parser.Parse(context);
            Assert.AreEqual(14, ss.Headers.Count);
            Assert.AreEqual("Sort Order", ss.Headers[0]);
            Assert.AreEqual("Sub Type", ss.Headers[4]);
            Assert.AreEqual(272, ss.Rows.Count);
            Assert.AreEqual("Kingdom of Bahrain", ss.Rows[12].Cells[2]);
            Assert.AreEqual(".bo", ss.Rows[20].Cells[13]);
         }
    }
}


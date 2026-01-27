using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
     [TestFixture]
    public class CsvParserTest
    {
         [Test]
         public void TestParseToxySpreadsheet()
         {
             string path = TestDataSample.GetFilePath("countrylist.csv", null);

            ParserContext context=new ParserContext(path);
            context.Properties.Add("ExtractHeader", "1");
            ISpreadsheetParser parser = (ISpreadsheetParser)ParserFactory.CreateSpreadsheet(context);
            ToxySpreadsheet ss= parser.Parse();
            ClassicAssert.AreEqual(1, ss.Tables.Count);
            ClassicAssert.AreEqual(14, ss.Tables[0].HeaderRows[0].Cells.Count);
            ClassicAssert.AreEqual("Sort Order", ss.Tables[0].HeaderRows[0].Cells[0].Value);
            ClassicAssert.AreEqual("Sub Type", ss.Tables[0].HeaderRows[0].Cells[4].Value);
            ClassicAssert.AreEqual(272, ss.Tables[0].Rows.Count);
            ClassicAssert.AreEqual(13, ss.Tables[0].Rows[12].RowIndex);
            ClassicAssert.AreEqual("Kingdom of Bahrain", ss.Tables[0].Rows[12].Cells[2].ToString());
            ClassicAssert.AreEqual(".bo", ss.Tables[0].Rows[20].Cells[13].ToString());

            ClassicAssert.AreEqual(272, ss.Tables[0].LastRowIndex);
         }
         [Test]
         public void TestParseIndexOutOfRange()
         {
             string path = TestDataSample.GetFilePath("countrylist.csv", null);
             ParserContext context = new ParserContext(path);
             context.Properties.Add("HasHeader", "1");
             ISpreadsheetParser parser = (ISpreadsheetParser)ParserFactory.CreateSpreadsheet(context);
             try
             {
                 ToxyTable ss = parser.Parse(1);
             }
             catch (ArgumentOutOfRangeException ex)
             {
                ClassicAssert.IsTrue(ex.Message.Contains("CSV only has one table"));
             }
         }
    }
}


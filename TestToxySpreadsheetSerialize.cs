using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class TestToxySpreadsheetSerialize
    {
        [Test]
        private void TestSpreadsheetSerialize(string filename = "countrylist.csv") // "Employee.xls"
        {
            ParserContext c = new ParserContext(TestDataSample.GetExcelPath(filename));
            // ParserContext c = new ParserContext(filename);
            var parser = ParserFactory.CreateSpreadsheet(c);
            var tss1 = parser.Parse();

            System.IO.MemoryStream ms = tss1.Serialize();


            ToxySpreadsheet tss2 = new ToxySpreadsheet();

            tss2.Deserialize(ms);



            comparespreadsheet(tss1, tss2);


        }

        private ToxySpreadsheet tssfromfile;
        private ToxySpreadsheet tssfromdb;

        public void comparespreadsheet(ToxySpreadsheet tss1, ToxySpreadsheet tss2)
        {


            //name
            Assert.AreEqual(tss1.Name, tss2.Name);


            //table
            Assert.AreEqual(tss1.Tables.Count, tss2.Tables.Count);

            for (int i = 0; i < tss1.Tables.Count; i++)
            {
                Assert.AreEqual(tss1.Tables[i].HasHeader, tss2.Tables[i].HasHeader);
                Assert.AreEqual(tss1.Tables[i].Name, tss2.Tables[i].Name);
                Assert.AreEqual(tss1.Tables[i].PageHeader, tss2.Tables[i].PageHeader);
                Assert.AreEqual(tss1.Tables[i].PageFooter, tss2.Tables[i].PageFooter);
                Assert.AreEqual(tss1.Tables[i].LastRowIndex, tss2.Tables[i].LastRowIndex);

                //mergecells
                Assert.AreEqual(tss1.Tables[i].MergeCells.Count, tss2.Tables[i].MergeCells.Count);
                for (int k = 0; k < tss1.Tables[i].MergeCells.Count; k++)
                {
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].FirstRow, tss2.Tables[i].MergeCells[k].FirstRow);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].LastRow, tss2.Tables[i].MergeCells[k].LastRow);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].FirstColumn, tss2.Tables[i].MergeCells[k].FirstColumn);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].LastColumn, tss2.Tables[i].MergeCells[k].LastColumn);

                }

                //columnheader
                if (tss1.Tables[i].HasHeader && tss2.Tables[i].HasHeader)
                {

                    Assert.AreEqual(tss1.Tables[i].ColumnHeaders.RowIndex, tss2.Tables[i].ColumnHeaders.RowIndex);
                    Assert.AreEqual(tss1.Tables[i].ColumnHeaders.LastCellIndex, tss2.Tables[i].ColumnHeaders.LastCellIndex);

                    //cells
                    Assert.AreEqual(tss1.Tables[i].ColumnHeaders.Cells.Count, tss2.Tables[i].ColumnHeaders.Cells.Count);
                    for (int j = 0; j < tss1.Tables[i].ColumnHeaders.Cells.Count; j++)
                    {
                        Assert.AreEqual(tss1.Tables[i].ColumnHeaders.Cells[j].Value, tss2.Tables[i].ColumnHeaders.Cells[j].Value);
                        Assert.AreEqual(tss1.Tables[i].ColumnHeaders.Cells[j].CellIndex, tss2.Tables[i].ColumnHeaders.Cells[j].CellIndex);
                        Assert.AreEqual(tss1.Tables[i].ColumnHeaders.Cells[j].Comment, tss2.Tables[i].ColumnHeaders.Cells[j].Comment);
                    }

                }

                //rows
                Assert.AreEqual(tss1.Tables[i].Rows.Count, tss2.Tables[i].Rows.Count);
                for (int m = 0; m < tss1.Tables[i].Rows.Count; m++)
                {
                    Assert.AreEqual(tss1.Tables[i].Rows[m].RowIndex, tss2.Tables[i].Rows[m].RowIndex);
                    Assert.AreEqual(tss1.Tables[i].Rows[m].LastCellIndex, tss2.Tables[i].Rows[m].LastCellIndex);

                    //cells
                    Assert.AreEqual(tss1.Tables[i].Rows[m].Cells.Count, tss2.Tables[i].Rows[m].Cells.Count);
                    for (int n = 0; n < tss1.Tables[i].Rows[m].Cells.Count; n++)
                    {
                        Assert.AreEqual(tss1.Tables[i].Rows[m].Cells[n].Value, tss2.Tables[i].Rows[m].Cells[n].Value);
                        Assert.AreEqual(tss1.Tables[i].Rows[m].Cells[n].CellIndex, tss2.Tables[i].Rows[m].Cells[n].CellIndex);
                        Assert.AreEqual(tss1.Tables[i].Rows[m].Cells[n].Comment, tss2.Tables[i].Rows[m].Cells[n].Comment);

                    }

                }




            }


        }


        public MemoryStream SpreadSheetSerialize(string filename)
        {

            ParserContext c = new ParserContext(TestDataSample.GetExcelPath(filename));
            // ParserContext c = new ParserContext(filename);
            var parser = ParserFactory.CreateSpreadsheet(c);
            ToxySpreadsheet tmpsheet =parser.Parse();
            MemoryStream ms =tmpsheet.Serialize();
            return ms;


        }

        public ToxySpreadsheet Spreadsheetfromfile(string fn)
        {
            ParserContext c = new ParserContext(TestDataSample.GetExcelPath(fn));
            // ParserContext c = new ParserContext(filename);
           
            var parser = ParserFactory.CreateSpreadsheet(c);
            ToxySpreadsheet tmpsheet = parser.Parse();
            return tmpsheet;


        }



        public ToxySpreadsheet SpreadSheetDeSerialize(MemoryStream ms)
        {
            ToxySpreadsheet tmpsheet = new ToxySpreadsheet();
            tmpsheet.Deserialize(ms);
            return tmpsheet;

        }

        public void comparefileanddb()
        {


            comparespreadsheet(tssfromdb, tssfromfile);


        }

        public List<string> GetAllWorkbook()
        {

            IEnumerable<string> dirs, files;
            List<string> lfiles = new List<string>();
            string subFolder = "Excel";

            string path = ConfigurationManager.AppSettings["testdataPath"].Replace('\\', Path.DirectorySeparatorChar);
            if (!path.EndsWith(string.Empty + Path.DirectorySeparatorChar))
            {
                path += Path.DirectorySeparatorChar;
            }


            //             dirs = Directory.EnumerateDirectories(path);
            //             foreach (string dir in dirs)
            //             {

            files = Directory.EnumerateFiles(path + subFolder);
            foreach (string excelfile in files)
            {

                lfiles.Add(excelfile);


            }
            /*            }*/


            return lfiles;

        }
        public void Testallworkbook()
        {




            List<string> files = GetAllWorkbook();
            foreach (string excelfile in files)
            {

                TestSpreadsheetSerialize(excelfile);


            }




        }


    }
}


using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestFixture]
    public class TestToxySpreadsheetSerialize
    {
        [Test]
        public void TestSpreadsheetSerialize()
        {
            ParserContext c = new ParserContext(TestDataSample.GetExcelPath("Employee.xls"));
            var parser = ParserFactory.CreateSpreadsheet(c);
            var tss1 = parser.Parse();

            System.IO.MemoryStream ms = tss1.Serialize();


            ToxySpreadsheet tss2 = new ToxySpreadsheet();

            tss2.Deserialize(ms);



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
                Assert.AreEqual(tss1.Tables[i].MergeCells.Count, tss1.Tables[i].MergeCells.Count);
                for (int k = 0; k < tss1.Tables[i].MergeCells.Count; k++)
                {
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].FirstRow, tss1.Tables[i].MergeCells[k].FirstRow);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].LastRow, tss1.Tables[i].MergeCells[k].LastRow);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].FirstColumn, tss1.Tables[i].MergeCells[k].FirstColumn);
                    Assert.AreEqual(tss1.Tables[i].MergeCells[k].LastColumn, tss1.Tables[i].MergeCells[k].LastColumn);

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




    }
}


using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    [TestClass]
    public class TestToxySpreadsheet
    {
        [TestMethod]
        public void TestToxyTableToDataTable()
        {
            #region create ToxyTable
            ToxyTable ttable = new ToxyTable();
            ttable.Name = "Test1";
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(0, "C1"));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(1, "C2"));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(2, "C3"));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(3, "C4"));
            ToxyRow trow1=new ToxyRow(0);
            trow1.Cells.Add(new ToxyCell(0,"1"));
            trow1.Cells.Add(new ToxyCell(1,"2"));
            trow1.Cells.Add(new ToxyCell(2,"3"));
            ttable.Rows.Add(trow1);

            ToxyRow trow2 = new ToxyRow(1);
            trow2.Cells.Add(new ToxyCell(0, "4"));
            trow2.Cells.Add(new ToxyCell(1, "5"));
            trow2.Cells.Add(new ToxyCell(3, "6"));
            trow2.LastCellIndex = 3;
            ttable.Rows.Add(trow2);

            ToxyRow trow3 = new ToxyRow(2);
            trow3.LastCellIndex = 3;
            trow3.Cells.Add(new ToxyCell(1, "7"));
            trow3.Cells.Add(new ToxyCell(2, "8"));
            trow3.Cells.Add(new ToxyCell(3, "9"));
            ttable.Rows.Add(trow3);

            ttable.LastColumnIndex = 3;
            #endregion
            DataTable dt = ttable.ToDataTable();
            Assert.AreEqual("Test1",dt.TableName);
            Assert.AreEqual(3, dt.Rows.Count);
            Assert.AreEqual(4, dt.Columns.Count);

            Assert.AreEqual("C1", dt.Columns[0].Caption);
            Assert.AreEqual("C2", dt.Columns[1].Caption);
            Assert.AreEqual("C3", dt.Columns[2].Caption);
            Assert.AreEqual("C4", dt.Columns[3].Caption);

            Assert.AreEqual("1", dt.Rows[0][0].ToString());
            Assert.AreEqual("2", dt.Rows[0][1].ToString());
            Assert.AreEqual("3", dt.Rows[0][2].ToString());
            Assert.AreEqual("4", dt.Rows[1][0].ToString());
            Assert.AreEqual("5", dt.Rows[1][1].ToString());
            Assert.AreEqual("6", dt.Rows[1][3].ToString());
            Assert.AreEqual("7", dt.Rows[2][1].ToString());
            Assert.AreEqual("8", dt.Rows[2][2].ToString());
            Assert.AreEqual("9", dt.Rows[2][3].ToString());
        }

        [TestMethod]
        public void TestToxyTableToDataTable_withEmptyColumnHeader()
        {
            #region create ToxyTable
            ToxyTable ttable = new ToxyTable();
            ttable.Name = "Test1";
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(0, "C1"));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(1, null));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(2, "C2"));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(3, null));
            ttable.ColumnHeaders.Cells.Add(new ToxyCell(4, "C4"));
            ToxyRow trow1 = new ToxyRow(0);
            trow1.Cells.Add(new ToxyCell(0, "1"));
            trow1.Cells.Add(new ToxyCell(1, "2"));
            trow1.Cells.Add(new ToxyCell(4, "3"));
            trow1.Cells.Add(new ToxyCell(5, "4"));
            trow1.LastCellIndex = 5;
            ttable.Rows.Add(trow1);

            ToxyRow trow2 = new ToxyRow(1);
            trow2.LastCellIndex = 3;
            trow2.Cells.Add(new ToxyCell(0, "5"));
            trow2.Cells.Add(new ToxyCell(1, "6"));
            trow2.Cells.Add(new ToxyCell(3, "7"));
            ttable.Rows.Add(trow2);

            ttable.LastColumnIndex = 5;
            #endregion

            DataTable dt = ttable.ToDataTable();
            Assert.AreEqual("Test1", dt.TableName);
            Assert.AreEqual(2, dt.Rows.Count);
            Assert.AreEqual(6, dt.Columns.Count);

            Assert.AreEqual("C1", dt.Columns[0].Caption);
            Assert.AreEqual("Column1", dt.Columns[1].Caption);
            Assert.AreEqual("C2", dt.Columns[2].Caption);
            Assert.AreEqual("Column2", dt.Columns[3].Caption);
            Assert.AreEqual("C4", dt.Columns[4].Caption);

            Assert.AreEqual("1", dt.Rows[0][0].ToString());
            Assert.AreEqual("2", dt.Rows[0][1].ToString());
            Assert.IsTrue(string.IsNullOrEmpty(dt.Rows[0][2].ToString()));
            Assert.IsTrue(string.IsNullOrEmpty(dt.Rows[0][3].ToString()));
            Assert.AreEqual("3", dt.Rows[0][4].ToString());
            Assert.AreEqual("4", dt.Rows[0][5].ToString());
            Assert.AreEqual("5", dt.Rows[1][0].ToString());
            Assert.AreEqual("6", dt.Rows[1][1].ToString());
            Assert.AreEqual("7", dt.Rows[1][3].ToString());
            Assert.IsTrue(string.IsNullOrEmpty(dt.Rows[1][2].ToString()));
            Assert.IsTrue(string.IsNullOrEmpty(dt.Rows[1][4].ToString()));
            Assert.IsTrue(string.IsNullOrEmpty(dt.Rows[1][5].ToString()));
        }
    }
}


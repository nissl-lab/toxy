using NUnit.Framework;
using NUnit.Framework.Legacy;
using System;
using System.IO;

namespace Toxy.Test
{
    [TestFixture]
    public class ColumnarSpreadsheetTest
    {
        // ── CSV ──────────────────────────────────────────────────────────────────

        [Test]
        public void TestCsvColumnarSpreadsheet_ColumnNamesAndRowCounts()
        {
            string path = TestDataSample.GetFilePath("countrylist.csv", null);
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            ClassicAssert.AreEqual(1, ss.Tables.Count);
            ToxyColumnarTable table = ss.Tables[0];

            // countrylist.csv has 14 header columns
            ClassicAssert.AreEqual(14, table.Columns.Count);
            ClassicAssert.AreEqual("Sort Order", table.Columns[0].Name);
            ClassicAssert.AreEqual("Sub Type", table.Columns[4].Name);

            // 272 data rows → 272 values per column
            ClassicAssert.AreEqual(272, table.Columns[0].Values.Count);
        }

        [Test]
        public void TestCsvColumnarSpreadsheet_ToParquet()
        {
            string path = TestDataSample.GetFilePath("countrylist.csv", null);
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".parquet");
            try
            {
                ss.ToParquet(tempFile);
                ClassicAssert.IsTrue(File.Exists(tempFile));
                ClassicAssert.Greater(new FileInfo(tempFile).Length, 0);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [Test]
        public void TestCsvColumnarSpreadsheet_ToArrow()
        {
            string path = TestDataSample.GetFilePath("countrylist.csv", null);
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".arrow");
            try
            {
                ss.ToArrow(tempFile);
                ClassicAssert.IsTrue(File.Exists(tempFile));
                ClassicAssert.Greater(new FileInfo(tempFile).Length, 0);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        // ── XLSX ─────────────────────────────────────────────────────────────────

        [Test]
        public void TestXlsxColumnarSpreadsheet_ColumnNamesAndRowCounts()
        {
            string path = TestDataSample.GetExcelPath("SheetWithColumnHeader.xlsx");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            ClassicAssert.IsTrue(ss.Tables.Count >= 1);
            ToxyColumnarTable table = ss.Tables[0];

            // SheetWithColumnHeader.xlsx has 4 columns: A, B, C, D
            ClassicAssert.AreEqual(4, table.Columns.Count);
            ClassicAssert.AreEqual("A", table.Columns[0].Name);
            ClassicAssert.AreEqual("B", table.Columns[1].Name);
            ClassicAssert.AreEqual("C", table.Columns[2].Name);
            ClassicAssert.AreEqual("D", table.Columns[3].Name);

            // 3 data rows
            ClassicAssert.AreEqual(3, table.Columns[0].Values.Count);
        }

        [Test]
        public void TestXlsxColumnarSpreadsheet_ToParquet()
        {
            string path = TestDataSample.GetExcelPath("SheetWithColumnHeader.xlsx");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".parquet");
            try
            {
                ss.ToParquet(tempFile);
                ClassicAssert.IsTrue(File.Exists(tempFile));
                ClassicAssert.Greater(new FileInfo(tempFile).Length, 0);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        [Test]
        public void TestXlsxColumnarSpreadsheet_ToArrow()
        {
            string path = TestDataSample.GetExcelPath("SheetWithColumnHeader.xlsx");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".arrow");
            try
            {
                ss.ToArrow(tempFile);
                ClassicAssert.IsTrue(File.Exists(tempFile));
                ClassicAssert.Greater(new FileInfo(tempFile).Length, 0);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
            }
        }

        // ── XLS ──────────────────────────────────────────────────────────────────

        [Test]
        public void TestXlsColumnarSpreadsheet_ParseSucceeds()
        {
            string path = TestDataSample.GetExcelPath("Formatting.xls");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            ClassicAssert.IsNotNull(ss);
            ClassicAssert.IsTrue(ss.Tables.Count >= 1);
            // Has at least one column
            ClassicAssert.IsTrue(ss.Tables[0].Columns.Count >= 1);
        }

        [Test]
        public void TestXlsColumnarSpreadsheet_ToParquet()
        {
            string path = TestDataSample.GetExcelPath("Formatting.xls");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".parquet");
            try
            {
                ss.ToParquet(tempFile);
                // For multi-table Excel, files use table-name suffixes
                // Check at least the first table file was created
                bool anyCreated = File.Exists(tempFile) ||
                    System.IO.Directory.GetFiles(
                        Path.GetDirectoryName(tempFile) ?? ".",
                        Path.GetFileNameWithoutExtension(tempFile) + "_*").Length > 0;
                ClassicAssert.IsTrue(anyCreated);
            }
            finally
            {
                // Clean up any generated files
                string dir = Path.GetDirectoryName(tempFile) ?? ".";
                string stem = Path.GetFileNameWithoutExtension(tempFile);
                foreach (var f in System.IO.Directory.GetFiles(dir, stem + "*"))
                    File.Delete(f);
            }
        }

        [Test]
        public void TestXlsColumnarSpreadsheet_ToArrow()
        {
            string path = TestDataSample.GetExcelPath("Formatting.xls");
            var context = new ParserContext(path);
            IColumnarSpreadsheetParser parser = ParserFactory.CreateColumnarSpreadsheet(context);
            ToxyColumnarSpreadsheet ss = parser.Parse();

            string tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".arrow");
            try
            {
                ss.ToArrow(tempFile);
                bool anyCreated = File.Exists(tempFile) ||
                    System.IO.Directory.GetFiles(
                        Path.GetDirectoryName(tempFile) ?? ".",
                        Path.GetFileNameWithoutExtension(tempFile) + "_*").Length > 0;
                ClassicAssert.IsTrue(anyCreated);
            }
            finally
            {
                string dir = Path.GetDirectoryName(tempFile) ?? ".";
                string stem = Path.GetFileNameWithoutExtension(tempFile);
                foreach (var f in System.IO.Directory.GetFiles(dir, stem + "*"))
                    File.Delete(f);
            }
        }

        // ── FromSpreadsheet ───────────────────────────────────────────────────────

        [Test]
        public void TestFromSpreadsheet_WithHeader()
        {
            var table = new ToxyTable();
            table.Name = "TestSheet";
            table.HeaderRows.Add(new ToxyRow(0));
            table.HeaderRows[0].Cells.Add(new ToxyCell(0, "Name"));
            table.HeaderRows[0].Cells.Add(new ToxyCell(1, "Age"));

            var row1 = new ToxyRow(1);
            row1.Cells.Add(new ToxyCell(0, "Alice"));
            row1.Cells.Add(new ToxyCell(1, "30"));
            row1.LastCellIndex = 1;
            table.Rows.Add(row1);

            var row2 = new ToxyRow(2);
            row2.Cells.Add(new ToxyCell(0, "Bob"));
            row2.Cells.Add(new ToxyCell(1, "25"));
            row2.LastCellIndex = 1;
            table.Rows.Add(row2);

            var ss = new ToxySpreadsheet();
            ss.Tables.Add(table);

            var columnar = ToxyColumnarSpreadsheet.FromSpreadsheet(ss);

            ClassicAssert.AreEqual(1, columnar.Tables.Count);
            ClassicAssert.AreEqual(2, columnar.Tables[0].Columns.Count);
            ClassicAssert.AreEqual("Name", columnar.Tables[0].Columns[0].Name);
            ClassicAssert.AreEqual("Age", columnar.Tables[0].Columns[1].Name);
            ClassicAssert.AreEqual(2, columnar.Tables[0].Columns[0].Values.Count);
            ClassicAssert.AreEqual("Alice", columnar.Tables[0].Columns[0].Values[0]);
            ClassicAssert.AreEqual("Bob", columnar.Tables[0].Columns[0].Values[1]);
            ClassicAssert.AreEqual("30", columnar.Tables[0].Columns[1].Values[0]);
            ClassicAssert.AreEqual("25", columnar.Tables[0].Columns[1].Values[1]);
        }

        [Test]
        public void TestFromSpreadsheet_WithoutHeader()
        {
            var table = new ToxyTable();
            table.Name = "NoHeader";
            var row1 = new ToxyRow(0);
            row1.Cells.Add(new ToxyCell(0, "X"));
            row1.Cells.Add(new ToxyCell(1, "Y"));
            row1.LastCellIndex = 1;
            table.Rows.Add(row1);

            var ss = new ToxySpreadsheet();
            ss.Tables.Add(table);

            var columnar = ToxyColumnarSpreadsheet.FromSpreadsheet(ss);

            ClassicAssert.AreEqual(2, columnar.Tables[0].Columns.Count);
            ClassicAssert.AreEqual("Column1", columnar.Tables[0].Columns[0].Name);
            ClassicAssert.AreEqual("Column2", columnar.Tables[0].Columns[1].Name);
            ClassicAssert.AreEqual("X", columnar.Tables[0].Columns[0].Values[0]);
            ClassicAssert.AreEqual("Y", columnar.Tables[0].Columns[1].Values[0]);
        }

        [Test]
        public void TestParserFactory_UnsupportedFormatThrows()
        {
            var context = new ParserContext("/tmp/test.pdf");
            Assert.Throws<NotSupportedException>(() => ParserFactory.CreateColumnarSpreadsheet(context));
        }
    }
}

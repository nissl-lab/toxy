using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using Apache.Arrow.Types;
using Parquet;
using Parquet.Data;
using Parquet.Schema;
// Resolve name conflicts between Apache.Arrow and Parquet.Schema
using ArrowField = Apache.Arrow.Field;
using ArrowSchema = Apache.Arrow.Schema;

namespace Toxy
{
    public class ToxyColumnarSpreadsheet
    {
        public string Name { get; set; }
        public List<ToxyColumnarTable> Tables { get; set; } = new List<ToxyColumnarTable>();

        /// <summary>
        /// Converts a row-oriented ToxySpreadsheet into a columnar representation.
        /// If a table has header rows, those are used as column names; otherwise
        /// column names are auto-generated as Column1, Column2, …
        /// </summary>
        public static ToxyColumnarSpreadsheet FromSpreadsheet(ToxySpreadsheet spreadsheet)
        {
            var columnar = new ToxyColumnarSpreadsheet();
            columnar.Name = spreadsheet.Name;

            foreach (var table in spreadsheet.Tables)
            {
                var columnarTable = new ToxyColumnarTable();
                columnarTable.Name = table.Name;

                // Determine column names and count
                int numCols;
                List<string> columnNames;

                if (table.HasHeader && table.HeaderRows.Count > 0)
                {
                    numCols = table.HeaderRows[0].Cells.Count;
                    columnNames = table.HeaderRows[0].Cells.Select(c => c.Value).ToList();
                }
                else
                {
                    numCols = table.Rows.Count > 0
                        ? table.Rows.Max(r => r.LastCellIndex + 1)
                        : 0;
                    columnNames = Enumerable.Range(0, numCols)
                        .Select(i => string.Format("Column{0}", i + 1))
                        .ToList();
                }

                // Create columns and fill values row by row
                var columns = columnNames.Select(name => new ToxyColumn { Name = name }).ToList();
                foreach (var row in table.Rows)
                {
                    for (int colIdx = 0; colIdx < columns.Count; colIdx++)
                    {
                        var cell = row.Cells.FirstOrDefault(c => c.CellIndex == colIdx);
                        columns[colIdx].Values.Add(cell != null ? cell.Value : string.Empty);
                    }
                }

                columnarTable.Columns = columns;
                columnar.Tables.Add(columnarTable);
            }

            return columnar;
        }

        /// <summary>
        /// Writes the spreadsheet to Parquet file(s). For a single table the file is written to
        /// <paramref name="outputPath"/>. For multiple tables one file per table is written
        /// (the table name is appended to the path stem).
        /// </summary>
        public void ToParquet(string outputPath)
        {
            if (Tables.Count == 0)
                return;

            if (Tables.Count == 1)
            {
                WriteTableToParquet(Tables[0], outputPath);
            }
            else
            {
                string dir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                string stem = Path.GetFileNameWithoutExtension(outputPath);
                string ext = Path.GetExtension(outputPath);
                foreach (var table in Tables)
                {
                    string tablePath = Path.Combine(dir, stem + "_" + SanitizeFileName(table.Name) + ext);
                    WriteTableToParquet(table, tablePath);
                }
            }
        }

        /// <summary>
        /// Writes the spreadsheet to Arrow IPC file(s). For a single table the file is written to
        /// <paramref name="outputPath"/>. For multiple tables one file per table is written.
        /// </summary>
        public void ToArrow(string outputPath)
        {
            if (Tables.Count == 0)
                return;

            if (Tables.Count == 1)
            {
                WriteTableToArrow(Tables[0], outputPath);
            }
            else
            {
                string dir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                string stem = Path.GetFileNameWithoutExtension(outputPath);
                string ext = Path.GetExtension(outputPath);
                foreach (var table in Tables)
                {
                    string tablePath = Path.Combine(dir, stem + "_" + SanitizeFileName(table.Name) + ext);
                    WriteTableToArrow(table, tablePath);
                }
            }
        }

        // ── Parquet helpers ──────────────────────────────────────────────────────

        private static void WriteTableToParquet(ToxyColumnarTable table, string path)
        {
            if (table.Columns.Count == 0)
            {
                File.WriteAllBytes(path, new byte[0]);
                return;
            }

            // Step 1: Build fields and collect typed arrays
            var fieldList = new List<Parquet.Schema.Field>(table.Columns.Count);
            var typedArrays = new List<System.Array>(table.Columns.Count);

            foreach (var col in table.Columns)
            {
                var result = CreateTypedParquetColumn(col.Name, col.Values);
                fieldList.Add(result.Item1);
                typedArrays.Add(result.Item2);
            }

            // Step 2: Create schema — this attaches the fields to the schema
            var schema = new ParquetSchema(fieldList);

            // Step 3: Get schema-attached DataField references and write
            DataField[] schemaFields = schema.GetDataFields();

            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            using var writer = ParquetWriter.CreateAsync(schema, fs).GetAwaiter().GetResult();
            using var rg = writer.CreateRowGroup();
            for (int i = 0; i < schemaFields.Length; i++)
            {
                rg.WriteColumnAsync(new DataColumn(schemaFields[i], typedArrays[i]))
                    .GetAwaiter().GetResult();
            }
        }

        private static Tuple<DataField, System.Array> CreateTypedParquetColumn(
            string name, List<string> rawValues)
        {
            // Try int
            if (rawValues.All(v => string.IsNullOrEmpty(v) || int.TryParse(v, out _)))
            {
                var field = new DataField<int?>(name);
                var values = rawValues
                    .Select(v => string.IsNullOrEmpty(v) ? (int?)null : int.Parse(v))
                    .ToArray();
                return Tuple.Create<DataField, System.Array>(field, values);
            }

            // Try double
            if (rawValues.All(v => string.IsNullOrEmpty(v) ||
                double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _)))
            {
                var field = new DataField<double?>(name);
                var values = rawValues
                    .Select(v => string.IsNullOrEmpty(v)
                        ? (double?)null
                        : double.Parse(v, CultureInfo.InvariantCulture))
                    .ToArray();
                return Tuple.Create<DataField, System.Array>(field, values);
            }

            // Try bool
            if (rawValues.All(v => string.IsNullOrEmpty(v) || bool.TryParse(v, out _)))
            {
                var field = new DataField<bool?>(name);
                var values = rawValues
                    .Select(v => string.IsNullOrEmpty(v) ? (bool?)null : bool.Parse(v))
                    .ToArray();
                return Tuple.Create<DataField, System.Array>(field, values);
            }

            // Try DateTime (stored as DateTimeOffset in Parquet)
            if (rawValues.All(v => string.IsNullOrEmpty(v) ||
                DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.None, out _)))
            {
                var field = new DataField<DateTimeOffset?>(name);
                var values = rawValues
                    .Select(v => string.IsNullOrEmpty(v)
                        ? (DateTimeOffset?)null
                        : (DateTimeOffset)DateTime.Parse(v, CultureInfo.InvariantCulture))
                    .ToArray();
                return Tuple.Create<DataField, System.Array>(field, values);
            }

            // Fall back to string
            {
                var field = new DataField<string>(name);
                var values = rawValues.Select(v => string.IsNullOrEmpty(v) ? null : v).ToArray();
                return Tuple.Create<DataField, System.Array>(field, values);
            }
        }

        // ── Arrow helpers ─────────────────────────────────────────────────────────

        private static void WriteTableToArrow(ToxyColumnarTable table, string path)
        {
            var schemaBuilder = new ArrowSchema.Builder();
            var arrays = new List<IArrowArray>(table.Columns.Count);

            foreach (var col in table.Columns)
            {
                var (array, arrowType) = BuildArrowArray(col.Values);
                schemaBuilder.Field(new ArrowField(col.Name, arrowType, nullable: true));
                arrays.Add(array);
            }

            var arrowSchema = schemaBuilder.Build();
            int rowCount = table.Columns.Count > 0 ? table.Columns[0].Values.Count : 0;
            var batch = new RecordBatch(arrowSchema, arrays, rowCount);

            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            using var writer = new ArrowFileWriter(fs, arrowSchema);
            writer.WriteRecordBatchAsync(batch).GetAwaiter().GetResult();
            writer.WriteEndAsync().GetAwaiter().GetResult();
        }

        private static (IArrowArray array, IArrowType type) BuildArrowArray(List<string> rawValues)
        {
            // Try int
            if (rawValues.All(v => string.IsNullOrEmpty(v) || int.TryParse(v, out _)))
            {
                var builder = new Int32Array.Builder();
                foreach (var v in rawValues)
                {
                    if (string.IsNullOrEmpty(v))
                        builder.AppendNull();
                    else
                        builder.Append(int.Parse(v));
                }
                return (builder.Build(), Int32Type.Default);
            }

            // Try double
            if (rawValues.All(v => string.IsNullOrEmpty(v) ||
                double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out _)))
            {
                var builder = new DoubleArray.Builder();
                foreach (var v in rawValues)
                {
                    if (string.IsNullOrEmpty(v))
                        builder.AppendNull();
                    else
                        builder.Append(double.Parse(v, CultureInfo.InvariantCulture));
                }
                return (builder.Build(), DoubleType.Default);
            }

            // Try bool
            if (rawValues.All(v => string.IsNullOrEmpty(v) || bool.TryParse(v, out _)))
            {
                var builder = new BooleanArray.Builder();
                foreach (var v in rawValues)
                {
                    if (string.IsNullOrEmpty(v))
                        builder.AppendNull();
                    else
                        builder.Append(bool.Parse(v));
                }
                return (builder.Build(), BooleanType.Default);
            }

            // Try DateTime → stored as Timestamp (milliseconds, UTC)
            if (rawValues.All(v => string.IsNullOrEmpty(v) ||
                DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.None, out _)))
            {
                var tsType = new TimestampType(TimeUnit.Millisecond, "UTC");
                var builder = new TimestampArray.Builder(tsType);
                foreach (var v in rawValues)
                {
                    if (string.IsNullOrEmpty(v))
                        builder.AppendNull();
                    else
                        builder.Append(
                            new DateTimeOffset(DateTime.Parse(v, CultureInfo.InvariantCulture)));
                }
                return (builder.Build(), tsType);
            }

            // Fall back to string
            {
                var builder = new StringArray.Builder();
                foreach (var v in rawValues)
                    builder.Append(v ?? string.Empty);
                return (builder.Build(), StringType.Default);
            }
        }

        // ── Utilities ─────────────────────────────────────────────────────────────

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "table";
            char[] invalid = Path.GetInvalidFileNameChars();
            char[] chars = name.Select(c => System.Array.IndexOf(invalid, c) >= 0 ? '_' : c).ToArray();
            return new string(chars);
        }
    }
}

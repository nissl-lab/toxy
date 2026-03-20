using CsvHelper;
using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class CSVSpreadsheetParser : ISpreadsheetParser
    {
		public ParserContext Context { get; set; }
		public event EventHandler<ParseSegmentEventArgs> ParseSegment;
		public CSVSpreadsheetParser(ParserContext context)
        {
            Context = context;
        }

        public class ParseSegmentEventArgs : EventArgs
        {
            public ParseSegmentEventArgs(string text, int rowIndex, int cellIndex)
            {
                RowIndex = rowIndex;
                CellIndex = cellIndex;
                Text = text;
            }
            public int RowIndex { get; set; }
            public int CellIndex { get; set; }
            public string Text { get; set; }
        }

        

        public ToxySpreadsheet Parse()
        {
            Utility.ValidateContext(Context);
            bool extractHeader = false;
            if (Context.Properties.ContainsKey("ExtractHeader"))
            {
                string sHasHeader = Context.Properties["ExtractHeader"].ToLower();
				if (sHasHeader == "1" || sHasHeader == "on" || sHasHeader == "true")
				{
					extractHeader = true;
				}
			}
            char delimiter =',';
            if (Context.Properties.ContainsKey("delimiter"))
            {
                delimiter = Context.Properties["delimiter"][0];
            }

            // the things we wanna close
            CsvReader reader = null;
			StreamReader sr = null;
            try
            {
				if (Context.IsStreamContext)
                {
                    sr = new StreamReader(Context.Stream, null, true, -1, true);
                }
                else
                {
					if (Context.Encoding == null)
                    {
                        sr = new StreamReader(Context.Path, true);
                    }
                    else
                    {
                        sr = new StreamReader(Context.Path, Context.Encoding, true);
                    }
                }
				// we do not want to close the Streams the User passed!
				reader = new CsvReader(sr, CultureInfo.InvariantCulture, Context.IsStreamContext);
                if (extractHeader)
                {
                    reader.Read();
                    reader.ReadHeader();
                }
                string[] headers = reader.HeaderRecord;
                ToxySpreadsheet ss = new ToxySpreadsheet();
                ToxyTable t1 = new ToxyTable();
                ss.Tables.Add(t1);

                int i = 0;
                if (headers != null && headers.Length > 0)
                {
                    t1.HeaderRows.Add(new ToxyRow(i));
                    i++;
                }
                for (int j = 0; j < headers.Length;j++ )
                {
                    t1.HeaderRows[0].Cells.Add(new ToxyCell(j, headers[j]));
                    t1.LastColumnIndex = t1.HeaderRows[0].Cells.Count-1;
                }
                while (reader.Read())
                {
                    ToxyRow tr = new ToxyRow(i);
                    tr.LastCellIndex = reader.Parser.Count -1;
                    if (tr.LastCellIndex > t1.LastColumnIndex)
                    {
                        t1.LastColumnIndex = tr.LastCellIndex;
                    }
                    tr.RowIndex = i;
                    for (int j = 0; j <= tr.LastCellIndex; j++)
                    {
                        if (ParseSegment != null)
                        {
							ParseSegment(this, new ParseSegmentEventArgs(reader[j], i, j));
                        }
                        ToxyCell c = new ToxyCell(j, reader[j]);
                        if (tr.LastCellIndex < c.CellIndex)
                        {
                            tr.LastCellIndex = c.CellIndex;
                        }
                        tr.Cells.Add(c);
                    }
                    
                    t1.Rows.Add(tr);
                    i++;
                }
                t1.LastRowIndex = i - 1;
                return ss;
            }
            finally
            {
                // will close the StreamReader and the Stream if we wanted so (not passed as Stream by the User see initialising)
                reader?.Dispose();
				// reader probably disposes it but if the reader could not br created it should be disposed!
				sr?.Dispose();
            }
        }

        public ToxyTable Parse(int sheetIndex)
        {
            if (sheetIndex > 0)
                throw new ArgumentOutOfRangeException("CSV only has one table");

            return Parse().Tables[0];
        }
    }
}

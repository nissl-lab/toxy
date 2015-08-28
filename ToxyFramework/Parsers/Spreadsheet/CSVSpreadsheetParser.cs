using LumenWorks.Framework.IO.Csv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class CSVSpreadsheetParser : ISpreadsheetParser
    {

        public CSVSpreadsheetParser(ParserContext context)
        {
            this.Context = context;
        }

        public class ParseSegmentEventArgs : EventArgs
        {
            public ParseSegmentEventArgs(string text, int rowIndex, int cellIndex)
            {
                this.RowIndex = rowIndex;
                this.CellIndex = cellIndex;
                this.Text = text;
            }
            public int RowIndex { get; set; }
            public int CellIndex { get; set; }
            public string Text { get; set; }
        }

        public event EventHandler<ParseSegmentEventArgs> ParseSegment;

        public ToxySpreadsheet Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            bool extractHeader=false;
            if (Context.Properties.ContainsKey("ExtractHeader"))
            {
                string sHasHeader = Context.Properties["ExtractHeader"].ToLower();
                if (sHasHeader == "1" || sHasHeader == "on" || sHasHeader == "true")
                    extractHeader = true;
            }
            char delimiter =',';
            if (Context.Properties.ContainsKey("delimiter"))
            {
                delimiter = Context.Properties["delimiter"][0];
            }

            
            Encoding encoding = Encoding.UTF8;

            StreamReader sr = null;
            try
            {
                if (Context.Encoding == null)
                {
                    sr = new StreamReader(Context.Path, true);
                }
                else
                {
                    sr = new StreamReader(Context.Path, true);
                }
                CsvReader reader=new CsvReader(sr, extractHeader,delimiter);
                string[] headers = reader.GetFieldHeaders();
                ToxySpreadsheet ss = new ToxySpreadsheet();
                ToxyTable t1 = new ToxyTable();
                ss.Tables.Add(t1);

                int i = 0;
                if (headers.Length > 0)
                {
                    t1.HeaderRows.Add(new ToxyRow(i));
                    i++;
                }
                for (int j = 0; j < headers.Length;j++ )
                {
                    t1.HeaderRows[0].Cells.Add(new ToxyCell(j, headers[j]));
                    t1.LastColumnIndex = t1.HeaderRows[0].Cells.Count-1;
                }
                while(reader.ReadNextRecord())
                {
                    ToxyRow tr=new ToxyRow(i);
                    tr.LastCellIndex = reader.FieldCount-1;
                    if (tr.LastCellIndex > t1.LastColumnIndex)
                    {
                        t1.LastColumnIndex = tr.LastCellIndex;
                    }
                    tr.RowIndex = i;
                    for (int j = 0; j <= tr.LastCellIndex; j++)
                    {
                        if (this.ParseSegment != null)
                        {
                            this.ParseSegment(this, new ParseSegmentEventArgs(reader[j], i, j));
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
                if (sr != null)
                    sr.Close();
            }
        }

        public ParserContext Context
        {
            get;
            set;
        }


        public ToxyTable Parse(int sheetIndex)
        {
            if (sheetIndex > 0)
                throw new ArgumentOutOfRangeException("CSV only has one table");

            return this.Parse().Tables[0];
        }
    }
}

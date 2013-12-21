using LumenWorks.Framework.IO.Csv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class CSVParser : PlainTextParser, ISpreadsheetParser
    {

        public CSVParser(ParserContext context):base(context)
        {
            this.Context = context;
        }

        public class ParseSegmentEventArgs : EventArgs
        {
            public ParseSegmentEventArgs(string text, int number)
            {
                this.LineNumber = number;
                this.Text = text;
            }
            public int LineNumber { get; set; }
            public string Text { get; set; }
        }

        public event EventHandler<ParseSegmentEventArgs> ParseSegment;

        public ToxySpreadsheet Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            bool hasHeader=false;
            string sHasHeader = Context.Properties["HasHeader"].ToLower();
            if (sHasHeader == "1" || sHasHeader == "on" || sHasHeader == "true")
                hasHeader = true;
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
                CsvReader reader=new CsvReader(sr, hasHeader,delimiter);
                string[] headers = reader.GetFieldHeaders();
                ToxySpreadsheet ss = new ToxySpreadsheet();
                ToxyTable t1 = new ToxyTable();
                ss.Tables.Add(t1);
                foreach (string header in headers)
                {
                    t1.ColumnHeaders.Add(header);
                }
                int i=0;
                while(reader.ReadNextRecord())
                {
                    ToxyRow tr=new ToxyRow();
                    tr.RowIndex = i;
                    for(int j=0;j<t1.ColumnHeaders.Count;j++)
                    {
                        ToxyCell c = new ToxyCell();
                        c.CellIndex = j;
                        c.Value = reader[j];
                        tr.Cells.Add(c);
                    }
                    
                    t1.Rows.Add(tr);
                }
                return ss;
            }
            finally
            {
                if (sr != null)
                    sr.Close();
            }
        }

        public override ParserContext Context
        {
            get;
            set;
        }
    }
}

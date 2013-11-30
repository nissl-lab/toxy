using LumenWorks.Framework.IO.Csv;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy.Parsers
{
    public class CSVParser : PlainTextParser, ISpreadsheetParser
    {
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

        public ToxySpreadsheet Parse(ParserContext context)
        {
            if (!File.Exists(context.Path))
                throw new FileNotFoundException("File " + context.Path + " is not found");

            bool hasHeader=false;
            if(context.Properties.ContainsKey("HasHeader"))
            {
                if(context.Properties["HasHeader"]=="1")
                    hasHeader=true;
            }
            char delimiter =',';
            if(context.Properties.ContainsKey("delimiter"))
            {
                delimiter = context.Properties["delimiter"][0];
            }

            
            Encoding encoding = Encoding.UTF8;

            StreamReader sr = null;
            try
            {
                if (context.Encoding == null)
                {
                    sr = new StreamReader(context.Path, true);
                }
                else
                {
                    sr = new StreamReader(context.Path, true);
                }
                CsvReader reader=new CsvReader(sr, hasHeader,delimiter);
                string[] headers = reader.GetFieldHeaders();
                ToxySpreadsheet ss = new ToxySpreadsheet();
                ss.Headers = new List<string>(headers.Length);
                foreach (string header in headers)
                {
                    ss.Headers.Add(header);
                }
                while(reader.ReadNextRecord())
                {
                    ToxyRow tr=new ToxyRow();
                    for(int j=0;j<ss.Headers.Count;j++)
                    {
                        tr.Cells.Add(reader[j]);
                    }
                    
                    ss.Rows.Add(tr);
                }
                return ss;
            }
            finally
            {
                if (sr != null)
                    sr.Close();
            }
        }
    }
}

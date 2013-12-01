using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Toxy.Parsers;

namespace Toxy
{
    public class ParserFactory
    {
        private ParserFactory() { }
        static Dictionary<string, Type> parserMapping = new Dictionary<string, Type>();

        static ParserFactory()
        {
            parserMapping.Add(".txt", typeof(PlainTextParser));
            parserMapping.Add(".xml", typeof(PlainTextParser));
            parserMapping.Add(".csv", typeof(CSVParser));
            parserMapping.Add(".xls", typeof(ExcelParser));
            parserMapping.Add(".xlsx", typeof(ExcelParser));
        }

        static string GetFileExtention(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (!parserMapping.ContainsKey(fi.Extension))
                throw new NotSupportedException("Extension " + fi.Extension + " is not supported");
            return fi.Extension;
        }
        public static ITextParser CreateText(string path)
        {
            string ext = GetFileExtention(path);
            Type parserType = parserMapping[ext];
            var obj = Activator.CreateInstance(parserType);
            if (!(obj is ITextParser))
                throw new InvalidDataException(ext+" is not supported by TextParser");
            ITextParser parser = (ITextParser)obj;
            return parser;
        }
        public static ISpreadsheetParser CreateSpreadsheet(string path)
        {
            string ext = GetFileExtention(path);
            Type parserType = parserMapping[ext];
            var obj = Activator.CreateInstance(parserType);
            if (!(obj is ISpreadsheetParser))
                throw new InvalidDataException(ext + " is not supported by SpreadsheetParser");
            ISpreadsheetParser parser = (ISpreadsheetParser)obj;
            return parser;            
        }
    }
}

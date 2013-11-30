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
            parserMapping.Add(".csv", typeof(CSVParser));
        }

        public static IParser Create(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (!parserMapping.ContainsKey(fi.Extension))
                throw new NotSupportedException("extension "+fi.Extension+" is not supported");

            Type parserType = parserMapping[fi.Extension];
            IParser parser = (IParser)Activator.CreateInstance(parserType);
            return parser;
        }


    }
}

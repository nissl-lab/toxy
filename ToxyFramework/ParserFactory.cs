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
        static Dictionary<string, List<Type>> parserMapping = new Dictionary<string, List<Type>>();

        static ParserFactory()
        {
            var typeTxt=new List<Type>();
            typeTxt.Add(typeof(PlainTextParser));
            parserMapping.Add(".txt", typeTxt);

            var typeXml = new List<Type>();
            typeXml.Add(typeof(PlainTextParser));
            parserMapping.Add(".xml", typeXml);

            var typeCSV = new List<Type>();
            typeCSV.Add(typeof(PlainTextParser));
            typeCSV.Add(typeof(CSVParser));
            parserMapping.Add(".csv", typeCSV);

            var typeXls = new List<Type>();
            typeXls.Add(typeof(ExcelParser));
            parserMapping.Add(".xls", typeXls);
            parserMapping.Add(".xlsx", typeXls);

            var typeDocx = new List<Type>();
            typeDocx.Add(typeof(WordParser));
            parserMapping.Add(".docx", typeDocx);
            
            var typeRtf = new List<Type>();
            typeRtf.Add(typeof(WordParser));
            parserMapping.Add(".rtf", typeRtf);

            var typeHtml = new List<Type>();
            typeHtml.Add(typeof(PlainTextParser));
            parserMapping.Add(".html", typeHtml);
            parserMapping.Add(".htm", typeHtml);

            var typeEml = new List<Type>();
            typeEml.Add(typeof(EMLParser));
            parserMapping.Add(".eml", typeEml);
        }

        static string GetFileExtention(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (!parserMapping.ContainsKey(fi.Extension))
                throw new NotSupportedException("Extension " + fi.Extension + " is not supported");
            return fi.Extension.ToLower();
        }
        static object CreateObject(ParserContext context, Type itype)
        {
             string ext = GetFileExtention(context.Path);
            var types= parserMapping[ext];
            object obj = null;
            bool isFound = false;
            foreach (Type type in types)
            {
                obj = Activator.CreateInstance(type, context);
                if (itype.IsAssignableFrom(obj.GetType()))
                {
                    isFound = true;
                    break;
                }
            }
            if (!isFound)
                throw new InvalidDataException(ext + " is not supported");
            return obj;
        }
        public static ITextParser CreateText(ParserContext context)
        {
            object obj = CreateObject(context, typeof(ITextParser));
            ITextParser parser = (ITextParser)obj;
            return parser;
        }
        public static ISpreadsheetParser CreateSpreadsheet(ParserContext context)
        {
            object obj = CreateObject(context, typeof(ISpreadsheetParser)); 
            ISpreadsheetParser parser = (ISpreadsheetParser)obj;
            return parser;
        }

        public static IDocumentParser CreateDocument(ParserContext context)
        {
            object obj = CreateObject(context, typeof(IDocumentParser));
            IDocumentParser parser = (IDocumentParser)obj;
            return parser;
        }
    }
}

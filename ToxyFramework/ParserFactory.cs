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
            var typeTxt = new List<Type>();
            typeTxt.Add(typeof(PlainTextParser));
            parserMapping.Add(".txt", typeTxt);

            var typeXml = new List<Type>();
            typeXml.Add(typeof(PlainTextParser));
            typeXml.Add(typeof(XMLDomParser));
            parserMapping.Add(".xml", typeXml);

            var typeCSV = new List<Type>();
            typeCSV.Add(typeof(PlainTextParser));
            typeCSV.Add(typeof(CSVSpreadsheetParser));
            parserMapping.Add(".csv", typeCSV);

            #region Office Formats
            var typeXls = new List<Type>();
            typeXls.Add(typeof(ExcelSpreadsheetParser));
            typeXls.Add(typeof(ExcelTextParser));
            typeXls.Add(typeof(OLE2MetadataParser));
            parserMapping.Add(".xls", typeXls);

            var typeXlsx = new List<Type>();
            typeXlsx.Add(typeof(ExcelSpreadsheetParser));
            typeXlsx.Add(typeof(ExcelTextParser));
            typeXlsx.Add(typeof(OOXMLMetadataParser));
            parserMapping.Add(".xlsx", typeXlsx);


            var typeOLE2 = new List<Type>();
            typeOLE2.Add(typeof(OLE2MetadataParser));
            parserMapping.Add(".ppt", typeOLE2);
            parserMapping.Add(".vsd", typeOLE2);
            parserMapping.Add(".pub", typeOLE2);
            parserMapping.Add(".shw", typeOLE2);
            parserMapping.Add(".sldprt", typeOLE2);

            var typePptx = new List<Type>();
            typePptx.Add(typeof(Powerpoint2007TextParser));
            typePptx.Add(typeof(Powerpoint2007SlideshowParser));
            typePptx.Add(typeof(OOXMLMetadataParser));
            parserMapping.Add(".pptx", typePptx);

            var typeOOXML = new List<Type>();
            typeOOXML.Add(typeof(OOXMLMetadataParser));
            parserMapping.Add(".pubx", typeOOXML);
            parserMapping.Add(".vsdx", typeOOXML);

            var typeDoc = new List<Type>();
            typeDoc.Add(typeof(Word2003DocumentParser));
            typeDoc.Add(typeof(Word2003TextParser));
            typeDoc.Add(typeof(OLE2MetadataParser));
            parserMapping.Add(".doc", typeDoc);

            var typeDocx = new List<Type>();
            typeDocx.Add(typeof(Word2007TextParser));
            typeDocx.Add(typeof(Word2007DocumentParser));
            typeDocx.Add(typeof(OOXMLMetadataParser));
            parserMapping.Add(".docx", typeDocx);
            #endregion

            var typeRtf = new List<Type>();
            typeRtf.Add(typeof(RTFTextParser));
            parserMapping.Add(".rtf", typeRtf);

            var typePdf = new List<Type>();
            typePdf.Add(typeof(PDFTextParser));
			typePdf.Add(typeof(PDFDocumentParser));
            parserMapping.Add(".pdf", typePdf);

            var typeHtml = new List<Type>();
            typeHtml.Add(typeof(PlainTextParser));
            typeHtml.Add(typeof(HtmlDomParser));
            parserMapping.Add(".html", typeHtml);
            parserMapping.Add(".htm", typeHtml);

            #region Email formats
            var typeEml = new List<Type>();
            typeEml.Add(typeof(EMLEmailParser));
            typeEml.Add(typeof(EMLTextParser));
            parserMapping.Add(".eml", typeEml);

            var typeCnm = new List<Type>();
            typeCnm.Add(typeof(CnmEmailParser));
            parserMapping.Add(".cnm", typeCnm);
            #endregion

            var typeVcard= new List<Type>();
            typeVcard.Add(typeof(VCardDocumentParser));
            typeVcard.Add(typeof(VCardTextParser));
            parserMapping.Add(".vcf", typeVcard);

            var typeZip = new List<Type>();
            typeZip.Add(typeof(ZipTextParser));
            parserMapping.Add(".zip", typeZip);

            var typeAudio = new List<Type>();
            typeAudio.Add(typeof(AudioMetadataParser));
            parserMapping.Add(".mp3", typeAudio);
            parserMapping.Add(".ape", typeAudio);
            parserMapping.Add(".wma", typeAudio);
            parserMapping.Add(".flac", typeAudio);
            parserMapping.Add(".aif", typeAudio);

            var typeImage = new List<Type>();
            typeImage.Add(typeof(ImageMetadataParser));
            parserMapping.Add(".jpeg", typeImage);
            parserMapping.Add(".jpg", typeImage);
            parserMapping.Add(".gif", typeImage);
            parserMapping.Add(".tiff", typeImage);
            parserMapping.Add(".png", typeImage);
        }

        static string GetFileExtention(string path)
        {
            FileInfo fi = new FileInfo(path);
            if (!parserMapping.ContainsKey(fi.Extension.ToLower()))
                throw new NotSupportedException("Extension " + fi.Extension + " is not supported");
            return fi.Extension.ToLower();
        }
        static object CreateObject(ParserContext context, Type itype, string operationName)
        {
            string ext = GetFileExtention(context.Path);
            var types = parserMapping[ext];
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
                throw new InvalidDataException(ext + " is not supported for " + operationName);
            return obj;
        }
        public static ITextParser CreateText(ParserContext context)
        {
            object obj = CreateObject(context, typeof(ITextParser), "CreateText");
            ITextParser parser = (ITextParser)obj;
            return parser;
        }
        public static IEmailParser CreateEmail(ParserContext context)
        {
            object obj = CreateObject(context, typeof(IEmailParser), "CreateEmail");
            IEmailParser parser = (IEmailParser)obj;
            return parser;
        }
        public static ISpreadsheetParser CreateSpreadsheet(ParserContext context)
        {
            object obj = CreateObject(context, typeof(ISpreadsheetParser), "CreateSpreadsheet");
            ISpreadsheetParser parser = (ISpreadsheetParser)obj;
            return parser;
        }

        public static IDocumentParser CreateDocument(ParserContext context)
        {
            object obj = CreateObject(context, typeof(IDocumentParser), "CreateDocument");
            IDocumentParser parser = (IDocumentParser)obj;
            return parser;
        }

        public static IDomParser CreateDom(ParserContext context)
        {
            object obj = CreateObject(context, typeof(IDomParser), "CreateDom");
            IDomParser parser = (IDomParser)obj;
            return parser;
        }

        public static VCardDocumentParser CreateVCard(ParserContext context)
        {
            object obj = CreateObject(context, typeof(VCardDocumentParser), "CreateVCard");
            VCardDocumentParser parser = (VCardDocumentParser)obj;
            return parser;
        }

        public static IMetadataParser CreateMetadata(ParserContext context)
        {
            object obj = CreateObject(context, typeof(IMetadataParser), "CreateMetadata");
            IMetadataParser parser = (IMetadataParser)obj;
            return parser;
        }

        public static ISlideshowParser CreateSlideshow(ParserContext context)
        {
            object obj = CreateObject(context, typeof(ISlideshowParser), "CreateSlideshow");
            ISlideshowParser parser = (ISlideshowParser)obj;
            return parser;            
        }
    }
}

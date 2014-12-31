using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    public class TestDataSample
    {
        public static string GetPdfPath(string filename)
        {
            return GetFilePath(filename, "Pdf");
        }
        public static string GetWordPath(string filename)
        {
            return GetFilePath(filename, "Word");
        }

        public static string GetRTFPath(string filename)
        {
            return GetFilePath(filename, "RTF");
        }
        public static string GetVCardPath(string filename)
        {
            return GetFilePath(filename, "Vcard");
        }

        public static string GetExcelPath(string filename)
        {
            return GetFilePath(filename, "Excel");
        }
        public static string GetTextPath(string filename)
        {
            return GetFilePath(filename, "txt");
        }
        public static string GetHtmlPath(string filename)
        {
            return GetFilePath(filename, "Html");
        }
        public static string GetAudioPath(string filename)
        {
            return GetFilePath(filename, "Audio");
        }
        public static string GetImagePath(string filename)
        {
            return GetFilePath(filename, "Image");
        }
        public static string GetOLE2Path(string filename)
        {
            return GetFilePath(filename, "OLE2");
        }
        public static string GetOOXMLPath(string filename)
        {
            return GetFilePath(filename, "OOXML");
        }
        public static string GetFilePath(string filename, string subFolder)
        {
            string path = ConfigurationManager.AppSettings["testdataPath"];
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }
            if(subFolder==null)
                return path + filename;
            else
                return path +subFolder+"\\" + filename;
        }
    }
}

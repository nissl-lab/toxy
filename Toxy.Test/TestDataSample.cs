using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.IO;

namespace Toxy.Test
{
    public class TestDataSample
    {
        const string samplePath = "..\\..\\..\\testdata\\";
        public static string GetPdfPath(string filename)
        {
            return GetFilePath(filename, "PDF");
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
            return GetFilePath(filename, "Txt");
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
            return GetFilePath(filename, "ole2");
        }
        public static string GetOOXMLPath(string filename)
        {
            return GetFilePath(filename, "ooxml");
        }
        public static string GetPowerpointPath(string filename)
        {
            return GetFilePath(filename, "powerpoint");
        }
        public static string GetCNMPath(string filename)
        {
            return GetFilePath(filename, "cnm");
        }
        public static string GetEmailPath(string filename)
        {
            return GetFilePath(filename, "email");
        }
        public static string GetFilePath(string filename, string subFolder)
        {
			string path = samplePath.Replace('\\',Path.DirectorySeparatorChar);
			if (!path.EndsWith(string.Empty+Path.DirectorySeparatorChar))
            {
				path += Path.DirectorySeparatorChar;
            }
            if(subFolder==null)
                return path + filename;
            else
				return path +subFolder+Path.DirectorySeparatorChar + filename;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace Toxy.Test
{
    public class TestDataSample
    {
        public static string GetFilePath(string filename)
        {
            string path = ConfigurationManager.AppSettings["testdataPath"];
            if (!path.EndsWith("\\"))
            {
                path += "\\";
            }
            return path + filename;
        }
    }
}

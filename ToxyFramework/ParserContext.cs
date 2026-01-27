using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ParserContext
    {
        public ParserContext(string path):this(path, null)
        { 
        
        }
        public ParserContext(string path, Encoding encoding)
        {
            this.Path = path;
            this.Properties = new Dictionary<string, string>();
            this.Encoding = encoding;
        }
        public string Path { get; set; }
        public Encoding Encoding { get; set; }
        
        public Dictionary<string, string> Properties { get; set; }
    }
}

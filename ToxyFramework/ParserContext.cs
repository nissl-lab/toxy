using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Math;
using FileSignatures;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy
{
    public class ParserContext
    {
        Stream _filestream;
        public ParserContext(string path):this(path, null)
        { 
        
        }
        public ParserContext(string path, Encoding encoding)
        {
            this.Path = path;
            this.Properties = new Dictionary<string, string>();
            this.Encoding = encoding;
        }
        public ParserContext(Stream stream):this(null, Encoding.UTF8)
        {
            this._filestream = stream;
        }
        
        public bool IsStreamContext { get=> this._filestream != null; }
        public Stream Stream { get => this._filestream; }
        public string Path { get; set; }

        public string Extension
        {
            get
            {
                return Utility.GetFileExtention(this);
            }
        }
        public string MimeType
        {
            get
            {
                FileFormat? fileformat = null;
                FileFormatInspector inspector = new FileFormatInspector();
                if (this.IsStreamContext)
                {
                    this.Stream.Position = 0;
                    fileformat = inspector.DetermineFileFormat(this.Stream);
                }
                else
                {
                    using var stream = File.OpenRead(this.Path);
                    fileformat = inspector.DetermineFileFormat(stream);
                }
                if (fileformat == null)
                    throw new InvalidDataException("File format could not be determined for the input stream");
                return fileformat.MediaType;

            }
        }
        public Encoding Encoding { get; set; }
        
        public Dictionary<string, string> Properties { get; set; }
    }
}

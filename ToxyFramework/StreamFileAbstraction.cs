using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Toxy
{
    internal class StreamFileAbstraction : TagLib.File.IFileAbstraction
    {
        public string Name { get; }
        public Stream ReadStream { get; }
        public Stream WriteStream { get; }

        public StreamFileAbstraction(string name, Stream stream)
        {
            Name = name;
            ReadStream = stream;
            WriteStream = stream;
        }

        public void CloseStream(Stream stream)
        {
            if (stream != null)
            {
                stream.Dispose();
            }
        }
    }
}

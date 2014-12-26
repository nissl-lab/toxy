using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
    public class ZipTextParser : PlainTextParser
    {
        public ZipTextParser(ParserContext context)
            : base(context)
        {
            this.Context = context;
        }
        public override string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            StringBuilder sb = new StringBuilder();
            using (ZipInputStream zipStream = new ZipInputStream(File.OpenRead(Context.Path)))
            {
                ZipEntry entry = zipStream.GetNextEntry();
                while (entry != null)
                {
                    sb.AppendLine(entry.Name);
                    entry = zipStream.GetNextEntry();
                }
            }
            return sb.ToString();
        }
    }
}

using ICSharpCode.SharpZipLib.Zip;
using System.IO;
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

            if (!Context.IsStreamContext&&!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            Stream stream = null;
            if (Context.IsStreamContext)
                stream = Context.Stream;
            else
                stream = File.OpenRead(Context.Path);
            StringBuilder sb = new StringBuilder();
            using (ZipInputStream zipStream = new ZipInputStream(stream))
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

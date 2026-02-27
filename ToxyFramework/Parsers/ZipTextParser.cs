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
            Utility.ValidateContext(Context);

            using var stream = Utility.GetStream(Context);
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

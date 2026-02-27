using ICSharpCode.SharpZipLib.Zip;
using PasswordProtectedChecker;
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

            if (!Context.IsStreamContext)
            {
                var checker = new Checker();
                if (checker.IsFileProtected(Context.Path).Protected)
                    throw new System.InvalidOperationException($"file {Context.Path} is encrypted");
            }

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

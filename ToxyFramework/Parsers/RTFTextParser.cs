using System.IO;

namespace Toxy.Parsers
{
    public class RTFTextParser : ITextParser
    {
        public RTFTextParser(ParserContext context)
        {
            Context = context;
        }
        public virtual ParserContext Context { get; set; }
        public string Parse()
        {
            if (!File.Exists(Context.Path))
            {
                throw new FileNotFoundException("File " + Context.Path + " is not found");
            }

            ReasonableRTF.RtfToTextConverter converter = new ReasonableRTF.RtfToTextConverter();
            ReasonableRTF.Models.RtfResult result = converter.Convert(File.ReadAllBytes(Context.Path));
            return result.Text;
        }
    }
}
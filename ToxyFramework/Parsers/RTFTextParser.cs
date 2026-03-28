using System;
using System.IO;
using Toxy.Base;

namespace Toxy.Parsers
{
    /// <summary>
    /// The <see cref="RTFTextParser"/> is used to extract the text of RTF Files/Streams
    /// </summary>
    public class RTFTextParser : BaseTextParser
    {
		/// <summary>
		/// Initializes the <see cref="RTFTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public RTFTextParser(ParserContext context) : base(context)
        { }
        internal override string ParseText(out IDisposable disposable)
        {
            Stream stream = Utility.GetStream(Context);
            disposable = Context.IsStreamContext ? null : stream;
            ReasonableRTF.RtfToTextConverter converter = new ReasonableRTF.RtfToTextConverter();
            ReasonableRTF.Models.RtfResult result = converter.Convert(stream);
            return result.Text;
        }
    }
}
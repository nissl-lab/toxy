using MimeKit;
using System.IO;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="EMLEmailParser"/> is used to extract EML Messages.
	/// </summary>
	public class EMLTextParser : MimeKitTextParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLTextParser(ParserContext context) : base(context)
		{ }

		private protected override MimeMessage GetMimeMessage(Stream stream)
		{
			return MimeMessage.Load(stream);
		}
	}
}

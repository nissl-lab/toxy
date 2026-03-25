using MimeKit;
using System.IO;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="EMLMetaParser"/> is used to get the metadata of an EML Message as <see cref="ToxyMetadata"/>.
	/// </summary>
	public class EMLMetaParser : MimeKitMetaParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLMetaParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLMetaParser(ParserContext context) : base(context)
		{ }

		private protected override MimeMessage GetMimeMessage(Stream stream)
		{
			return MimeMessage.Load(stream);
		}
	}
}

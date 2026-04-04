using MimeKit;
using System;
using System.IO;
using Toxy.Base;
using Toxy.Helpers;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="EMLMetaParser"/> is used to get the metadata of an EML Message as <see cref="ToxyMetadata"/>.
	/// </summary>
	public class EMLMetaParser : BaseMetaParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLMetaParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLMetaParser(ParserContext context) : base(context)
		{ }

		internal override ToxyMetadata ParseMeta(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = MimeMessage.Load(stream))
			{
				return MimeMessageHelper.ConvertToMeta(message);
			}
		}
	}
}

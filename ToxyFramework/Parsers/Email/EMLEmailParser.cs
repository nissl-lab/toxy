using MimeKit;
using System;
using System.IO;
using Toxy.Base;
using Toxy.Helpers;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="EMLEmailParser"/> is used to convert an EML Message to a <see cref="ToxyEmail"/>.
	/// </summary>
	public class EMLEmailParser : BaseEmailParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLEmailParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLEmailParser(ParserContext context) : base(context)
		{ }

		internal override ToxyEmail ParseEmail(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = MimeMessage.Load(stream))
			{
				return MimeMessageHelper.ConvertToToxyEmail(message);
			}
		}
	}
}
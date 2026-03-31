using MimeKit;
using System;
using System.IO;
using Toxy.Base;
using Toxy.Helpers;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="EMLTextParser"/> is used to extract the Text of EML Messages.
	/// </summary>
	public class EMLTextParser : BaseTextParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLTextParser(ParserContext context) : base(context)
		{ }

		internal override string ParseText(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = MimeMessage.Load(stream))
			{
				return MimeMessageHelper.ConvertToText(message);
			}
		}
	}
}

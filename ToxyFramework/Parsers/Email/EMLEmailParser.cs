using HLIB.MailFormats;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using Toxy.Parsers.Email.Base;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="MimeKitEmailParser"/> is used to convert an EML Message to a <see cref="ToxyEmail"/>.
	/// </summary>
	public class EMLEmailParser : MimeKitEmailParser
	{
		/// <summary>
		/// Initializes the <see cref="EMLEmailParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public EMLEmailParser(ParserContext context) : base(context)
		{ }

		private protected override MimeMessage GetMimeMessage(Stream stream)
		{
			return MimeMessage.Load(stream);
		}
	}
}

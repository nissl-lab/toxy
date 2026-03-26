using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using Toxy.Base;

namespace Toxy.Parsers.Email.Base
{
	// This is a Base class since it will make it much easier to support MBOX Files.
	// To support MBOX a sub class need to be created and the "GetMimeMessage" needs to be overwritten with:
	// MimeParser parser = new MimeParser(stream, MimeFormat.Mbox);
	// MimeMessage message = parser.ParseMessage();
	/// <summary>
	/// The <see cref="MimeKitEmailParser"/> is used to gets the <see cref="MimeMessage"/> as <see cref="ToxyEmail"/>.
	/// </summary>
	public abstract class MimeKitEmailParser : BaseEmailParser
	{
		/// <summary>
		/// Initializes the <see cref="MimeKitTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public MimeKitEmailParser(ParserContext context) : base(context)
		{ }

#nullable enable
		/// <summary>
		/// Gets the <paramref name="addresses"/> as <see cref="List{string}"/>, which contains each address as a <see cref="string"/>.
		/// </summary>
		/// <param name="addresses">The <see cref="InternetAddressList"/> of an Email.</param>
		/// <returns>Returns the <paramref name="addresses"/> as strings.</returns>
		private static List<string> ToStringAdressList(InternetAddressList? addresses)
		{
			if (addresses is null) return new List<string>();

			List<string> result = new List<string>(addresses.Count);
			foreach (InternetAddress address in addresses)
			{
				result.Add(address.ToString());
			}
			return result;
		}
#nullable restore

		internal override sealed ToxyEmail ParseEmail(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = GetMimeMessage(stream))
			{
				ToxyEmail email = new ToxyEmail();
				email.From = message.From.ToString();
				email.To = ToStringAdressList(message.To);
				email.Cc = ToStringAdressList(message.Cc);
				email.TextBody = message.TextBody;
				email.HtmlBody = message.HtmlBody;
				email.Subject = message.Subject;
				email.ArrivalTime = message.Date.DateTime;
				return email;
			}
		}
		/// <summary>
		/// Gets the <see cref="MimeMessage"/> of the Parser.
		/// This Method will be overriden by every Parser, which uses the <see cref="MimeMessage"/> to extract Text.
		/// </summary>
		/// <param name="stream">The <see cref="Stream"/> of the Mail File.</param>
		/// <returns>Returns the <see cref="MimeMessage"/> of the Parser.</returns>
		private protected abstract MimeMessage GetMimeMessage(Stream stream);
	}
}

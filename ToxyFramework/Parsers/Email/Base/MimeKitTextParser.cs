using HtmlAgilityPack;
using MimeKit;
using System;
using System.IO;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	// This is a Base class since it will make it much easier to support MBOX Files.
	// To support MBOX a sub class need to be created and the "GetMimeMessage" needs to be overwritten with:
	// MimeParser parser = new MimeParser(stream, MimeFormat.Mbox);
	// MimeMessage message = await parser.ParseMessageAsync();
	/// <summary>
	/// The <see cref="MimeKitTextParser"/> is used to extract the Text of a <see cref="MimeMessage"/>
	/// </summary>
	public abstract class MimeKitTextParser : BaseTextParser
	{
		/// <summary>
		/// Initializes the <see cref="MimeKitTextParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public MimeKitTextParser(ParserContext context) : base(context)
		{ }

		internal override sealed string ParseText(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = GetMimeMessage(stream))
			{

				StringBuilder sb = new StringBuilder();
				TryAppend(ref sb, "[From] ", message.From.ToString());
				TryAppend(ref sb, "[To] ", message.To.ToString());
				TryAppend(ref sb, "[CC] ", message.Cc.ToString());
				TryAppend(ref sb, "[Subject] ", message.Subject);
				sb.AppendLine();

				if (!string.IsNullOrEmpty(message.TextBody))
				{
					sb.Append(message.TextBody);
				}
				else if (!string.IsNullOrEmpty(message.HtmlBody))
				{
					HtmlDocument document = new HtmlDocument();
					document.LoadHtml(message.HtmlBody);
					sb.AppendLine(document.DocumentNode.InnerText);
				}
				return sb.ToString();
			}
		}

		/// <summary>
		/// Appends the <paramref name="content"/> to the <paramref name="sb"/> if it is not empty (<see cref="string.IsNullOrEmpty(string)"/>).
		/// </summary>
		/// <param name="sb">The <see cref="StringBuilder"/> to append the <paramref name="prefix"/> &amp; <paramref name="content"/> to.</param>
		/// <param name="prefix">The Prefix, which will be appened before the <paramref name="content"/>.</param>
		/// <param name="content">The actual content.</param>
		/// <param name="newLine">Specifies if a new Line should be addeed at the end.</param>
		/// <returns>Returns <see langword="true"/> if the <paramref name="content"/> has been appended otherwise <see langword="false"/> will be returned.</returns>
		private static bool TryAppend(ref StringBuilder sb, string prefix, string content, bool newLine = true)
		{
			if (string.IsNullOrEmpty(content)) return false;

			sb.Append(prefix);
			if (newLine)
			{
				sb.AppendLine(content);
			}
			else
			{
				sb.Append(content);
			}
			return true;
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

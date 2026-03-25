using DocumentFormat.OpenXml.InkML;
using HtmlAgilityPack;
using MimeKit;
using System;
using System.IO;
using System.Text;
using Toxy.Base;

namespace Toxy.Parsers
{
	/// <summary>
	/// The <see cref="MimeKitMetaParser"/> is used to extract the <see cref="ToxyMetadata"/> of a <see cref="MimeMessage"/>
	/// </summary>
	public abstract class MimeKitMetaParser : BaseMetaParser
	{
		/// <summary>
		/// Initializes the <see cref="MimeKitMetaParser"/>
		/// </summary>
		/// <param name="context">The <see cref="ParserContext"/> of the Parser.</param>
		public MimeKitMetaParser(ParserContext context) : base(context)
		{ }

		internal override sealed ToxyMetadata ParseMeta(out IDisposable disposable)
		{
			Stream stream = Utility.GetStream(Context);
			disposable = Context.IsStreamContext ? null : stream;
			using (MimeMessage message = GetMimeMessage(stream))
			{
				ToxyMetadata meta = new ToxyMetadata();
				meta.Add(nameof(message.Bcc), message.Bcc);
				meta.Add(nameof(message.Cc), message.Cc);
				meta.Add(nameof(message.Date), message.Date);
				meta.Add(nameof(message.From), message.From);
				meta.Add(nameof(message.Importance), message.Importance);
				meta.Add(nameof(message.InReplyTo), message.InReplyTo);
				meta.Add(nameof(message.MessageId), message.MessageId);
				meta.Add(nameof(message.MimeVersion), message.MimeVersion);
				meta.Add(nameof(message.Priority), message.Priority);
				meta.Add(nameof(message.ReplyTo), message.ReplyTo);
				meta.Add(nameof(message.ResentBcc), message.ResentBcc);
				meta.Add(nameof(message.ResentCc), message.ResentCc);
				meta.Add(nameof(message.ResentDate), message.ResentDate);
				meta.Add(nameof(message.ResentFrom), message.ResentFrom);
				meta.Add(nameof(message.ResentMessageId), message.ResentMessageId);
				meta.Add(nameof(message.ResentReplyTo), message.ResentReplyTo);
				meta.Add(nameof(message.ResentSender), message.ResentSender);
				meta.Add(nameof(message.ResentTo), message.ResentTo);
				meta.Add(nameof(message.Sender), message.Sender);
				meta.Add(nameof(message.Subject), message.Subject);
				meta.Add(nameof(message.To), message.To);
				meta.Add(nameof(message.XPriority), message.XPriority);
				return meta;
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

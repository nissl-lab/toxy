using HtmlAgilityPack;
using MimeKit;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Helpers
{
	/// <summary>
	/// Contains Methods converting a <see cref="MimeMessage"/>
	/// </summary>
	internal static class MimeMessageHelper
	{
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

		/// <summary>
		/// Converts the <paramref name="message"/> to a <see cref="ToxyEmail"/>.
		/// </summary>
		/// <param name="message">The <see cref="MimeMessage"/>, which should be converted.</param>
		/// <returns>Returns the converted <see cref="ToxyEmail"/>.</returns>
		internal static ToxyEmail ConvertToToxyEmail(MimeMessage message)
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
		/// <summary>
		/// Converts the <paramref name="message"/> to a <see cref="ToxyMetadata"/>.
		/// It will extract only Metadata.
		/// </summary>
		/// <param name="message">The <see cref="MimeMessage"/>, which should be converted.</param>
		/// <returns>Returns the converted <see cref="ToxyMetadata"/>.</returns>
		internal static ToxyMetadata ConvertToMeta(MimeMessage message)
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
		/// <summary>
		/// Converts the <paramref name="message"/> to plain Text.
		/// </summary>
		/// <param name="message">The <see cref="MimeMessage"/>, which should be converted.</param>
		/// <returns>Returns the plain Text and some additional Informations of the <paramref name="message"/>.</returns>
		internal static string ConvertToText(MimeMessage message)
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
	}
}

using System;
using System.Text;
using System.Text.RegularExpressions;

using InSolve.dmach.Mail.Decoders;

namespace InSolve.dmach.Mail
{
	/// <summary>
	/// Вспомогательные методы для парсера.
	/// </summary>
	public static class ParserTools
	{
		#region Regular expression fields

		/// <summary>
		/// Заголовок
		/// </summary>
		internal static Regex _regexHeader = new Regex(@"\A[a-z\-]+(?=:)", RegexOptions.Compiled | RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase);
		/// <summary>
		/// Выделяет разделитель (используется группа "bound")
		/// </summary>
		private static Regex _regexBound = new Regex("(?:boundary=)(?:\\\")?(?<bound>[a-z0-9\\-=_\\.]+)(?:\\\")?", RegexOptions.Compiled | RegexOptions.IgnoreCase );

		#endregion

		public static string ConvertMimeHeader(string header, out bool correctHeader, Action<ParserError, string> errorHandler)
		{
			correctHeader = _regexHeader.IsMatch(header);

			if (header.LastIndexOf("?=") == -1)
				return header;

			Match match = Mime.RegexMime.Match(header);
			if (!match.Success)
				return header;
			
			StringBuilder buffer = new StringBuilder(header);
			while(match.Success)
			{
				bool confidently;
				
				DecoderBase decoder = DecoderSelector.GetHeaderDecoder(
					match.Groups["transfer"].Value,
					out confidently
					);

                if (!confidently && errorHandler != null)
                    errorHandler(ParserError.FailGetDecoderFromMimeString, match.Groups["transfer"].Value);
			
				decoder.Add(match.Groups["data"].Value);

				Encoding encoding;
				try
				{
					encoding = Encoding.GetEncoding(match.Groups["enc"].Value);
				}
				catch(Exception exc)
				{
                    if (errorHandler != null)
                        errorHandler(ParserError.UnknownEncoding, "'" + match.Groups["enc"].Value + "': " + exc.Message);

					encoding = Encoding.Default;
				}

				string data = encoding.GetString(decoder.ToArray());

				buffer.Remove(match.Index, match.Length);
				buffer.Insert(match.Index, data);

				match = Mime.RegexMime.Match(buffer.ToString());
			}

			return buffer.ToString();
		}

		
		public static void ParseHeader(string str, HeaderCollection headers, IMessageConstructor constructor, Action<ParserError, string> errorHandler, bool upperLevel)
		{
			Match match = _regexHeader.Match(str);
			if (!match.Success)
			{
				if (errorHandler != null)
					errorHandler(ParserError.HeaderMismatchMimePattern, str);

				return;
			}

			string name		= match.Value.ToLower();
			string value	= str.Substring(match.Length + 1).Trim();

			headers.Add(name, value);
			if (upperLevel)
				constructor.AddHeader(name, value);

			return;
		}

	
		public static string GetNextBound(HeaderCollection headers)
		{
			string content_type = headers.Get(Mime.Content_Type);
			if (content_type == null)
				return null;

			Match match = _regexBound.Match(content_type);

			return (match.Success)? "--" + match.Groups["bound"].Value : null;
		}
	
	
		public static byte[] GetBytes(string str)
		{
			if (str == null)
				return null;

			byte[] b = new byte[str.Length];
			for (int i = 0; i < str.Length; i++)
				b[i] = (byte)str[i];

			return b;
		}

        public static bool EqualsIgnoreCase(this string s, string value)
        {
            return s.Equals(value, StringComparison.CurrentCultureIgnoreCase);
        }
	}
}

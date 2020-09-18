using System;
using System.Text;
using System.Collections;
using System.Runtime.Serialization;
using System.Collections.Specialized;
using System.Text.RegularExpressions;

namespace InSolve.dmach.Mail
{
	// Суперважный класс! Один из самых основных.

	/// <summary>
	/// Коллекция заголовков.
	/// </summary>
	[Serializable]
	public class HeaderCollection :
		NameValueCollection
	{
		#region Regular Expression Fields

		/// <summary>
		/// Выделяет имя файла (испльзуется группа "file")
		/// </summary>
		private static Regex _regexFile = new Regex("(?:(?:file)?name=)(?:[\\\"'])?(?<file>[^\\\"']+)(?:[\\\"'])?", RegexOptions.Compiled | RegexOptions.IgnoreCase);
		/// <summary>
		/// Выделяет кодировку (используется группа "enc")
		/// </summary>
		private static Regex _regexCharset = new Regex("(?:charset=)(?:\\\")?(?<enc>[a-z0-9\\-]+)(?:\\\")?", RegexOptions.Compiled | RegexOptions.IgnoreCase );
		/// <summary>
		/// Выделяет тип (группа "type") и подтип (группа "subtype")
		/// </summary>
		private static Regex _regexType = new Regex(@"(?<type>[a-z0-9\-]+)/(?<subtype>[a-z0-9\-]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase );

		#endregion

		#region Collection Methods

		public override void Add(string name, string value)
		{
			if (name == null)
				throw new ArgumentNullException("name");

			base.Add (name.Trim().ToLower(), value);
		}


		public void Replace(string name, string value)
		{
			if (name == null)
				throw new ArgumentNullException("name");

			name = name.Trim().ToLower();

			Remove(name);
			base.Add(name, value);
		}

		public bool Contains(string name)
		{
			if (name == null)
				throw new ArgumentNullException("name");

			name = name.Trim().ToLower();

			return GetValues(name) != null;
		}

		public string GetLast(string name)
		{
			if (name != null)
				name = name.Trim().ToLower();

			string[] values = GetValues(name);

			return (values != null)? values[values.Length - 1] : null;
		}

		public string GetFirst(string name)
		{
			if (name != null)
				name = name.Trim().ToLower();

			string[] values = GetValues(name);

			return (values != null)? values[0] : null;
		}

		#endregion

		#region Mime properties

		public string MimeContentID
		{
			get
			{
				if (Contains("message-id"))
					return Get("message-id");
				
				return Contains(Mime.Content_ID)? Get(Mime.Content_ID) : string.Empty;
			}
		}

		public string MimeContentDescription
		{
			get
			{
				return Contains(Mime.Content_Description)? Get(Mime.Content_Description) : string.Empty;
			}
		}

		public string MimeContentName
		{
			get
			{
				if (Contains(Mime.Content_Type))
				{
					Match match = _regexFile.Match(Get(Mime.Content_Type));
					if (match.Success)
						return match.Groups["file"].Value;
				}

				if (Contains(Mime.Content_Disposition))
				{
					Match match = _regexFile.Match(Get(Mime.Content_Disposition));
					if (match.Success)
						return match.Groups["file"].Value;
				}

				return string.Empty;
			}
		}

		public string MimeContentType
		{
			get
			{
				string content_type = Get(Mime.Content_Type);
				if (content_type == null)
					return "";

				Match match = _regexType.Match(content_type);
				return (match.Success)? match.Groups["type"].Value : "";
			}
		}

		public string MimeContentSubtype
		{
			get
			{
				string content_type = Get(Mime.Content_Type);
				if (content_type == null)
					return "";

				Match match = _regexType.Match(content_type);
				return (match.Success)? match.Groups["subtype"].Value : "";
			}
		}

		#endregion

		#region Methods

        public Encoding GetEncoding()
        {
            return GetEncoding(Encoding.Default);
        }

        public Encoding GetEncoding(int defaultCodePage)
        {
            return GetEncoding(Encoding.GetEncoding(defaultCodePage));
        }

		public Encoding GetEncoding(Encoding defaultEnc)
		{
            if (defaultEnc == null)
                throw new ArgumentNullException("defaultEnc");

			string content_type = Get(Mime.Content_Type);
			if (content_type == null)
				return defaultEnc;

			Match match = _regexCharset.Match(content_type);
			if (!match.Success)
                return defaultEnc;

			try
			{
				return Encoding.GetEncoding(match.Groups["enc"].Value);
			}
			catch
			{
				return defaultEnc;
			}
		}

        public AttachmentFormat GetAttachmentFormat(out string ContentType)
        {
            ContentType = Get(Mime.Content_Type);
            AttachmentFormat format = AttachmentFormat.Unknown;
            string type = MimeContentType;

            if (type.EqualsIgnoreCase("text") || type == "")
            {
                string subtype = MimeContentSubtype;
                if (subtype.EqualsIgnoreCase("plain") || subtype == "")
                    format = AttachmentFormat.PlainText;
                else if (subtype.EqualsIgnoreCase("html"))
                    format = AttachmentFormat.Html;
            }
            else if (type.EqualsIgnoreCase("message"))
                format = AttachmentFormat.ByteStreamMessage;
            else
                format = AttachmentFormat.ByteStream;

            return format;
        }

        public AttachmentFormat GetAttachmentFormat()
        {
            string dummy;
            return GetAttachmentFormat(out dummy);
        }

		#endregion
	}

}

using System;
using System.IO;

using InSolve.dmach.Mail.Decoders;

namespace InSolve.dmach.Mail
{
	/// <summary>
	/// Базовый класс для прикрепленных элементов
	/// </summary>
	[Serializable]
	public class Attachment :
		IAttachmentConstructor
	{
		#region Fields
		
		/// <summary>
		/// Описание контента
		/// </summary>
		protected HeaderCollection _headers;

		private Stream _stream;
		
		#endregion

		#region Constructors

        public Attachment(HeaderCollection headers, Stream stream, DecoderSelector decoders, bool appendBodyData, Action<ParserError, string> errorHandler)
        {
            if (headers == null)
                throw new ArgumentNullException("headers");
            if (decoders == null)
                throw new ArgumentNullException("decoders");

            _headers = headers;

            if (appendBodyData && stream != null)
                _stream = new StreamConverter(stream, decoders.GetDecoder(headers, errorHandler));
        }

		#endregion

		#region Properties

		/// <summary>
		/// Заголовки
		/// </summary>
		public HeaderCollection Headers
		{
			get
			{
				return _headers;
			}
		}

		protected Stream InnerStream
		{
			get
			{
				if (_stream is StreamConverter)
					throw new InvalidOperationException("IAttachmentConstructor cannot close");

				return _stream;
			}
		}

		public string Name
		{
			get
			{
				return _headers.MimeContentName;
			}
		}

        public FileInfo TagFile { get; set; }

		#endregion

		#region IAttachmentConstructor Members

		System.IO.Stream IAttachmentConstructor.Stream
		{
			get
			{
				if (_stream != null && !(_stream is StreamConverter))
                    throw new InvalidOperationException("Attachment has been closed");

				return _stream;
			}
		}

		void IAttachmentConstructor.CompleteConstruction()
		{
			if (_stream != null && _stream is StreamConverter)
			{
				_stream.Flush();
				_stream = (_stream as StreamConverter).GetTargetStream();
			}
		}

		#endregion

		#region Boolean properties

		public bool IsContainsAttachmentData
		{
			get
			{
				return (_stream != null);
			}
		}

		public bool IsAttachmentMessageBody
		{
			get
			{
                AttachmentFormat format = _headers.GetAttachmentFormat();
				return (format == AttachmentFormat.Html || format == AttachmentFormat.PlainText);
			}
		}

        public bool IsAttachmentMessageBodyHtml
        {
            get
            {
                return (_headers.GetAttachmentFormat() == AttachmentFormat.Html);
            }
        }

        public bool IsAttachmentMessageBodyTextPlain
        {
            get
            {
                return (_headers.GetAttachmentFormat() == AttachmentFormat.PlainText);
            }
        }

        public bool IsAttachmentMessageStream
        {
            get
            {
                return (_headers.GetAttachmentFormat() == AttachmentFormat.ByteStreamMessage);
            }
        }

        public bool IsAttachmentFile
		{
			get
			{
                return (_headers.GetAttachmentFormat() == AttachmentFormat.ByteStream);
			}
		}

		#endregion

		#region Methods

        public void CloseInnerStream()
        {
            if (_stream != null)
                _stream.Close();
        }

		public byte[] ToArray()
		{
			return StreamTools.GetBytesFromBegin(InnerStream);
		}

        public override string ToString()
        {
            string text = null;
            switch (_headers.GetAttachmentFormat())
            {
                case AttachmentFormat.PlainText:
                case AttachmentFormat.Html:
                    text = _headers.GetEncoding().GetString(ToArray());
                    break;
                case AttachmentFormat.ByteStream:
                    text = _headers.MimeContentName;
                    break;
                case AttachmentFormat.ByteStreamMessage:
                    text = "Included Message";
                    break;
                case AttachmentFormat.Unknown:
                default:
                    text = "Unknown attachment format";
                    break;
            }

            return text;
        }

		#endregion
	}
}

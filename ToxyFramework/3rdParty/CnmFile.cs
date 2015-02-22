using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using InSolve.dmach.Mail;
using InSolve.dmach.Mail.Decoders;

namespace CnmViewer
{
    public class CnmFile : IMessageConstructor
    {
        public CnmFile(FileInfo source)
        {
            DateTime dt;
            try
            {
                using (StreamReader r =  source.OpenText())
                {
                    for (int i = 0; ; i++)
                    {
                        if (i > 3)
                            throw new ApplicationException("Can't found time");

                        string str = r.ReadLine();
                        int semicolon = str.IndexOf(';');
                        if (semicolon != -1)
                        {
                            if ((semicolon + 1) == str.Length)
                                continue;
                            else
                                str = str.Substring(semicolon + 1);
                        }
                        str = str.Trim();
                        if (DateTime.TryParse(str, out dt))
                            break;
                    }
                }
            }
            catch
            {
                dt = DateTime.MinValue;
            }

            Time = dt;
            File = source;
        }

        public FileInfo File { get; private set; }
        public DateTime Time { get; private set; }

        public Stream TextHtml { get; set; }
        public string TextPlain { get; set; }
        public string Headers { get; set; }
        public List<Attachment> Files { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Errors { get; set; }

        StringBuilder _headersBuffer;
        StringBuilder _errorsBuffer;

        public bool HasErrors
        {
            get
            {
                return (Errors != null);
            }
        }

        public bool IsParsed
        {
            get
            {
                return (Headers != null);
            }
        }

        public void Parse()
        {
            if (File != null)
                using (FileStream s = File.OpenRead())
                    Parser.Parse(this, s);
        }

        #region IMessageConstructor Members

        public void AddHeader(string name, string value)
        {
            if (_headersBuffer == null)
                _headersBuffer = new StringBuilder();

            _headersBuffer.AppendLine(string.Format("{0}: {1}", name, value));

            if (name.EqualsIgnoreCase("from"))
                From = value.Trim();
            else if (name.EqualsIgnoreCase("to"))
                To = value.Trim();
            else if (name.EqualsIgnoreCase("subject"))
                Subject = value.Trim();
        }

        public IAttachmentConstructor GetAttachmentConstructor(HeaderCollection headers)
        {
            AttachmentFormat format = headers.GetAttachmentFormat();
            bool requiredBody = false;
            Stream stream = null;

            if (format == AttachmentFormat.Html)
            {
                requiredBody = true;
                stream = new MemoryStream();
            }
            else if (format == AttachmentFormat.PlainText)
            {
                requiredBody = true;
                stream = new MemoryStream();
            }

            Attachment att = new Attachment(
                headers, 
                stream,
                new DecoderSelector(), 
                requiredBody, 
                (err, comm) => ErrorHandler(err, comm));

            return att;
        }

        public void CompleteAttachment(IAttachmentConstructor att)
        {
            Attachment a = (Attachment)att;
            try
            {
                if (a.IsAttachmentMessageBody)
                {
                    if (a.IsAttachmentMessageBodyHtml)
                        TextHtml = new MemoryStream(a.ToArray());
                    else if (a.IsAttachmentMessageBodyTextPlain)
                        TextPlain = a.Headers.GetEncoding(1251).GetString(a.ToArray()).Trim();
                }
                else
                {
                    if (!string.IsNullOrEmpty(a.Headers.MimeContentName))
                    {
                        if (Files == null)
                            Files = new List<Attachment>();

                        Files.Add(a);
                    }
                }
            }
            finally
            {
                a.CloseInnerStream();
            }
        }

        public void CompleteConstruction()
        {
            if (_headersBuffer != null)
            {
                Headers = _headersBuffer.ToString().TrimEnd();
                _headersBuffer = null;
            }
            if (_errorsBuffer != null)
            {
                Errors = _errorsBuffer.ToString().TrimEnd();
                _errorsBuffer = null;
            }
        }

        public void ErrorHandler(ParserError err, string comment)
        {
            if (_errorsBuffer == null)
                _errorsBuffer = new StringBuilder();

            _errorsBuffer.AppendLine(string.Format("{0}: {1}", err, comment));
        }

        #endregion
    }
}

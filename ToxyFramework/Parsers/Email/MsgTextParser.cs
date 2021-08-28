using MsgReader;
using MsgReader.Outlook;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace Toxy.Parsers
{
    /// <summary>
    /// msg file parser
    /// </summary>
    /// <remarks>
    /// http://www.codeproject.com/Articles/19571/MsgReader-DLL
    /// </remarks>
    public class MsgTextParser:PlainTextParser
    {
        public MsgTextParser(ParserContext context)
            : base(context)
        {
            this.Context = context;
        }

        public override string Parse()
        {
                StringBuilder result = new StringBuilder();

                using (var stream = File.OpenRead(this.Context.Path))
                using (var reader = new Storage.Message(stream))
                {
                    if (reader.Sender != null)
                    {
                        result.AppendFormat("[From] {0}{1}", string.IsNullOrEmpty(reader.Sender.DisplayName)
                            ? reader.Sender.Email :
                            string.Format("{0}<{1}>", reader.Sender.DisplayName, reader.Sender.Email), Environment.NewLine);
                    }
                    if (reader.Recipients.Count > 0)
                    {
                        StringBuilder recipientTo = new StringBuilder();
                        StringBuilder recipientCc = new StringBuilder();
                        StringBuilder recipientBcc = new StringBuilder();
                        foreach (var recipient in reader.Recipients)
                        {
                            string sRecipient = null;
                            if (string.IsNullOrEmpty(recipient.DisplayName))
                            {
                                sRecipient = recipient.Email+";";
                            }
                            else
                            {
                                sRecipient = string.Format("{0}<{1}>;", recipient.DisplayName, recipient.Email);
                            }

                            if (recipient.Type == RecipientType.To)
                            {
                                recipientTo.Append(sRecipient);
                            }
                            else if (recipient.Type == RecipientType.Cc)
                            {
                                recipientCc.Append(sRecipient);
                            }
                            else if (recipient.Type == RecipientType.Bcc)
                            {
                                recipientBcc.Append(sRecipient);
                            }
                        }
                        if (recipientTo.Length > 0)
                        {
                            result.Append("[To] ");
                            result.AppendLine(recipientTo.ToString());
                        }
                        if (recipientCc.Length > 0)
                        {
                            result.Append("[Cc] ");
                            result.AppendLine(recipientCc.ToString());
                        }
                        if (recipientBcc.Length>0)
                        {
                            result.Append("[Bcc] ");
                            result.AppendLine(recipientBcc.ToString());
                        }

                    }
                    if (!string.IsNullOrEmpty(reader.Subject))
                        result.AppendFormat("[Subject] {0}{1}", reader.Subject, Environment.NewLine);

                    result.AppendLine();
                    result.AppendLine(reader.BodyText);

                }
            
                return result.ToString();
        }
    }
}

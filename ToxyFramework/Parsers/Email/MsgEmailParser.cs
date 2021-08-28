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
    /// msg file parser for ToxyEmail structure
    /// </summary>
    public class MsgEmailParser:IEmailParser
    {
        public MsgEmailParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyEmail Parse()
        {
            ToxyEmail result = new ToxyEmail();
           using (var stream = File.OpenRead(this.Context.Path))
           using (var reader = new Storage.Message(stream))
           {
               if (reader.Sender != null)
               {
                   result.From = string.IsNullOrEmpty(reader.Sender.DisplayName)
                            ? reader.Sender.Email :
                            string.Format("{0}<{1}>", reader.Sender.DisplayName, reader.Sender.Email);
               }
               if (reader.Recipients.Count > 0)
               {
                   foreach (var recipient in reader.Recipients)
                   {
                       string sRecipient = null;
                       if (string.IsNullOrEmpty(recipient.DisplayName))
                       {
                           sRecipient = recipient.Email;
                       }
                       else
                       {
                           sRecipient = string.Format("{0}<{1}>", recipient.DisplayName, recipient.Email);
                       }
                       if (recipient.Type == RecipientType.To)
                       {
                           result.To.Add(sRecipient);
                       }
                       else if (recipient.Type == RecipientType.Cc)
                       {
                           result.Cc.Add(sRecipient);
                       }
                       else if (recipient.Type ==RecipientType.Bcc)
                       {
                           result.Bcc.Add(sRecipient);
                       }
                   }
               }
               result.Subject = reader.Subject;
               result.TextBody = reader.BodyText;
               result.HtmlBody = reader.BodyHtml;
           }
           return result;
        }

        public ParserContext Context
        {
            get;
            set;
        }
    }
}

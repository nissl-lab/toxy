using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using VCardReader;

namespace Toxy.Parsers
{
    public class VCardTextParser : PlainTextParser
    {
        public VCardTextParser(ParserContext context):base(context)
        {
        
        }
        public override string Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            string path = Context.Path;
            StringBuilder sb = new StringBuilder();
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    var card = new VCard(sr);

                    if (!string.IsNullOrEmpty(card.FormattedName))
                        sb.AppendFormat("[Full Name]{0}" + Environment.NewLine, card.FormattedName);
                    if (!string.IsNullOrEmpty(card.GivenName))
                        sb.AppendFormat("[First Name]{0}" + Environment.NewLine, card.GivenName);
                    if (!string.IsNullOrEmpty(card.AdditionalNames))
                        sb.AppendFormat("[Middle Name]{0}" + Environment.NewLine, card.AdditionalNames);
                    if (!string.IsNullOrEmpty(card.FamilyName))
                        sb.AppendFormat("[Last Name]{0}" + Environment.NewLine, card.FamilyName);
                    if (!string.IsNullOrEmpty(card.ProductId))
                        sb.AppendFormat("[Product ID]{0}" + Environment.NewLine, card.ProductId);
                    if (!string.IsNullOrEmpty(card.Organization))
                        sb.AppendFormat("[Orgnization]{0}" + Environment.NewLine, card.Organization);
                    if (card.Sources.Count > 0)
                    {
                        sb.AppendLine("[Sources]");
                        foreach (var vSource in card.Sources)
                        {
                            sb.AppendLine(vSource.Uri.OriginalString);
                        }
                    }
                    if (!string.IsNullOrEmpty(card.Title))
                        sb.AppendFormat("[Title]{0}" + Environment.NewLine, card.Title);
                    if (card.Gender != Gender.Unknown)
                        sb.AppendFormat("[Gender]{0}" + Environment.NewLine, card.Gender);
                    if (card.Nicknames.Count > 0)
                    {
                        sb.AppendFormat("[Nickname]{0}" + Environment.NewLine, card.Nicknames[0]);
                    }
                    if (card.DeliveryAddresses.Count > 0)
                    {
                        sb.AppendLine("[Addresses]");
                        foreach (var dAddr in card.DeliveryAddresses)
                        {
                            sb.Append(dAddr.AddressType + ":");
                            if (!string.IsNullOrEmpty(dAddr.Street))
                                sb.Append(dAddr.Street + ",");
                            if (!string.IsNullOrEmpty(dAddr.City))
                                sb.Append(dAddr.City + ",");
                            if (!string.IsNullOrEmpty(dAddr.Region))
                                sb.Append(dAddr.Region + ",");
                            if (!string.IsNullOrEmpty(dAddr.Country))
                                sb.Append(dAddr.Country + ",");

                            sb.AppendLine();
                        }
                    }
                    if (card.Phones.Count > 0)
                    {
                        sb.AppendLine("[Phones]");
                        foreach (var vphone in card.Phones)
                        {
                            sb.AppendFormat("{0}:{1}" + Environment.NewLine, vphone.PhoneType, vphone.FullNumber);
                        }
                    }
                    if (card.EmailAddresses.Count > 0)
                    {
                        sb.AppendLine("[Emails]");
                        foreach (var vEmail in card.EmailAddresses)
                        {
                            sb.AppendFormat("{0}:{1}" + Environment.NewLine, vEmail.EmailType, vEmail.Address);
                        }
                    }
                    if (card.Websites.Count > 0)
                    {
                        sb.AppendLine("[Websites]");
                        foreach (var vWebsite in card.Websites)
                        {
                            sb.AppendLine(vWebsite.Url);
                        }
                    }
                    sb.AppendLine();
                }
            }
            return sb.ToString();
        }
    }
}

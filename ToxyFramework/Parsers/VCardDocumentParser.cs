using System.IO;
using VCardReader;

namespace Toxy.Parsers
{
    public class VCardDocumentParser
    {
        public ParserContext Context
        {
            get;
            set;
        }

        public VCardDocumentParser(ParserContext context)
        {
            this.Context = context;
        }

        public ToxyBusinessCards Parse()
        {
            if (!File.Exists(Context.Path))
                throw new FileNotFoundException("File " + Context.Path + " is not found");

            string path = Context.Path;
            ToxyBusinessCards tbcs = new ToxyBusinessCards();
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    var card = new VCard(sr);
                    ToxyBusinessCard tbc = new ToxyBusinessCard();
                    tbc.Name = new ToxyName();
                    if (!string.IsNullOrEmpty(card.FormattedName))
                        tbc.Name.FullName = card.FormattedName;

                    tbc.Name.FirstName = card.GivenName;
                    tbc.Name.MiddleName = card.AdditionalNames;
                    tbc.Name.LastName = card.FamilyName;
                    tbc.ProductID = card.ProductId;
                    foreach(var vSource in card.Sources)
                    {
                        tbc.Sources.Add(vSource.Uri.OriginalString);
                    }
                    tbc.Orgnization = card.Organization;
                    tbc.Title = card.Title;
                    tbc.Gender = card.Gender == Gender.Male ? GenderType.Male : GenderType.Female;
                    if (card.Nicknames.Count > 0)
                    {
                        tbc.NickName = new ToxyName();
                        tbc.NickName.FullName = card.Nicknames[0];
                    }
                    foreach (var dAddr in card.DeliveryAddresses)
                    {
                        var tAddr= new ToxyAddress();
                        tAddr.City = dAddr.City;
                        tAddr.Street = dAddr.Street;
                        tAddr.Country = dAddr.Country;
                        tAddr.Region = dAddr.Region;
                        tAddr.PostalCode = dAddr.PostalCode;
                        tbc.Addresses.Add(tAddr);
                    }

                    foreach (var vphone in card.Phones)
                    {
                        tbc.Contacts.Add(new ToxyContact(vphone.PhoneType.ToString(), vphone.FullNumber)); 
                    }
                    foreach (var vEmail in card.EmailAddresses)
                    {
                        tbc.Contacts.Add(new ToxyContact(vEmail.EmailType.ToString(), vEmail.Address)); 
                    }
                    foreach (var vWebsite in card.Websites)
                    {
                        tbc.Contacts.Add(new ToxyContact("Url-"+ vWebsite.WebsiteType.ToString(), vWebsite.Url));
                    }

                    tbcs.Cards.Add(tbc);
                }
            }
            return tbcs;
        }
    }
}

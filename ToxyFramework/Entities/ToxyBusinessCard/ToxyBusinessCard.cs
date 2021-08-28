using System;
using System.Collections.Generic;
using System.Text;
using VCardReader;

namespace Toxy
{


    public enum GenderType
    { 
        Male,
        Female
    }


    public class ToxyBusinessCard
    {
        public ToxyBusinessCard()
        {
            this.Contacts = new List<ToxyContact>();
            this.Addresses = new List<ToxyAddress>();
            this.Sources = new List<string>();
            this.Photos = new List<string>();
        }
        public ToxyName Name { get; set; }
        public ToxyName NickName { get; set; }
        public string Orgnization { get; set; }
        public string Title { get; set; }
        public string Class { get; set; }
        public string TimeZone { get; set; }
        public string UID { get; set; }
        public string GEO { get; set; }
        public string Label { get; set; }
        public List<string> Categories { get; set; }
        public DateTime Birthday { get; set; }
        public List<string> Photos {get;set;}
        public List<string> Sources { get; set; }
        public string ProductID { get; set; }
        public List<ToxyAddress> Addresses { get; set; }
        public List<ToxyContact> Contacts { get; set; }

        public GenderType Gender { get; set; }

        public VCard ToVCard()
        {
            var card = new VCard();
            if(this.Name!=null)
            {
            card.DisplayName = this.Name.FullName;
            card.FamilyName = this.Name.LastName;
            card.GivenName = this.Name.FirstName;
            }
            card.Gender = this.Gender== GenderType.Male?VCardReader.Gender.Male: VCardReader.Gender.Female;
            if (this.NickName != null)
                card.Nicknames.Add(this.NickName.FullName);

            card.Organization = this.Orgnization;
            if (this.Contacts != null)
            {
                foreach (var contact in this.Contacts)
                {
                    card.Phones.Add(new Phone(contact.Name, PhoneTypes.Cellular));
                }
            }
            return card;
        }

        public override string ToString()
        {
            return string.Format("{0}-{1}",this.Name.FullName,this.Title);
        }
    }



}


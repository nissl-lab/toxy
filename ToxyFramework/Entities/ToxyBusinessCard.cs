using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyBusinessCards
    {
        public ToxyBusinessCards()
        {
            this.Cards = new List<ToxyBusinessCard>();
        }
        public List<ToxyBusinessCard> Cards { get; set; }
    }

    public enum GenderType
    { 
        Male,
        Female
    }

    public class ToxyName
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string FullName { get; set; }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(FullName))
                return FullName;

            if (string.IsNullOrEmpty(MiddleName))
                return FirstName + " " + LastName;

            return FirstName + " " + MiddleName +" " + LastName;
        }
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
    }



}


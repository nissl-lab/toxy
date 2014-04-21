using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyContact
    {
        public static string CELLPHONE="Cellular";
        public static string HOME = "Home";
        public static string WORK = "Work";
        public static string FAX = "Fax";
        public static string INTERNET = "Internet";
        public static string DEFAULT = "Default";
        public static string URL_DEFAULT = "Url-Default";
        public static string VOICE = "Voice";
        

        public ToxyContact(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }
        public ToxyContact(string name, string value, string note)
            : this(name, value)
        {
            this.Note = note;
        }
        public string Name { get; set; }
        public string Value { get; set; }
        public string Note { get; set; }

        public override string ToString()
        {
            return string.Format("{0}:{1}", this.Name,this.Value);
        }
    }
}

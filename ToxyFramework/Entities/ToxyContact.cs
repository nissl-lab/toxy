using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy.Entities
{
    public class ToxyContact
    {
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
    }
}

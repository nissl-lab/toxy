using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
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

            return FirstName + " " + MiddleName + " " + LastName;
        }
    }
}

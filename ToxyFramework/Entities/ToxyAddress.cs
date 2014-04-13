using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public class ToxyAddress
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }

        public override string ToString()
        {
            string result = string.Empty;

            if (this.Street != null)
            {
                result = this.Street;
            }
            if (this.City != null)
            {
                if (result.Length > 0)
                {
                    result += ";";
                }
                result += this.City;
            }
            if (this.Region != null)
            {
                if (result.Length > 0)
                {
                    result += ";";
                }
                result += this.Region;
            }
            if (this.PostalCode != null)
            {
                if (result.Length > 0)
                {
                    result += ";";
                }
                result += this.PostalCode;
            }
            if (this.Country != null)
            {
                if (result.Length > 0)
                {
                    result += ";";
                }
                result += this.Country;
            }
            return result;
        }
    }
}

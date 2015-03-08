using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public class ToxySlide
    {
        public ToxySlide()
        {
            this.Texts = new List<string>();
        }
        public List<string> Texts { get; set; }
        public string Note { get; set; }
    }
}

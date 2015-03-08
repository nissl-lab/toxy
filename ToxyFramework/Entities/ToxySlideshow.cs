using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public class ToxySlideshow
    {
        public ToxySlideshow()
        {
            this.Slides = new List<ToxySlide>();
        }

        public List<ToxySlide> Slides { get; set; }
    }
}

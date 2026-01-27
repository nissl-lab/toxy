using System.Collections.Generic;

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

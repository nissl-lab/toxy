using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public interface ISlideshowParser
    {
        ToxySlideshow Parse();
        ToxySlide Parse(int slideIndex);
        ParserContext Context { get; set; }
    }
}

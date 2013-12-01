using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IToxyDocument
    {
        string Title { get; set; }
        List<IParagraph> Paragraphs { get; set; }
        int TotalPageNumber { get; set; }
        List<Property> Properties { get; set; }
    }
}

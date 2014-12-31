using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Toxy
{
    public interface IMetadataParser
    {
        ToxyMetadata Parse();
        ParserContext Context { get; set; }
    }
}

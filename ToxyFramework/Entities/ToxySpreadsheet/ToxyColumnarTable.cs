using System.Collections.Generic;

namespace Toxy
{
    public class ToxyColumnarTable
    {
        public string Name { get; set; }
        public List<ToxyColumn> Columns { get; set; } = new List<ToxyColumn>();
    }
}

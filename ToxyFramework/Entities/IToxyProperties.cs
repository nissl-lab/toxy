using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public interface IToxyProperties
    {
        Dictionary<string, object> Properties { get; set; }
    }
}

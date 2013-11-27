using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    interface IProperty
    {
        string Key { get; set; }
        string Value { get; set; }
    }
}

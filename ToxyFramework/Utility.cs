using System;
using System.Collections.Generic;
using System.Text;

namespace Toxy
{
    public static class Utility
    {
        public static bool IsTrue(string sValue)
        {
            sValue = sValue.ToLower();
            if (sValue == "1" || sValue == "on" || sValue == "true")
                return true;
            else
                return false;
        }
    }
}

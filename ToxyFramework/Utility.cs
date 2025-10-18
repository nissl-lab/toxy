using System;

namespace Toxy
{
    public static class Utility
    {
        public static bool IsTrue(string sValue)
        {
            return (sValue.Equals("1", StringComparison.OrdinalIgnoreCase)) || (sValue.Equals("on", StringComparison.OrdinalIgnoreCase)) || (sValue.Equals("true", StringComparison.OrdinalIgnoreCase));
        }
    }
}

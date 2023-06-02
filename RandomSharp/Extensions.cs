using System;

namespace RandomSharp
{
    public static class Extensions
    {
        public static decimal ToDecimal(this double d)
        {
            return Convert.ToDecimal(d);
        }

        public static int ToInt(this double d)
        {
            return Convert.ToInt32(d);
        }
    }
}

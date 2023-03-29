﻿using System;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    public static class IComparableExtension
    {
        public static bool InRange<T>(this T value, T from, T to) where T : IComparable<T>
        {
            return value.CompareTo(from) >= 1 && value.CompareTo(to) <= -1;
        }
    }
}

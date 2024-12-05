using System;
using System.Collections.Generic;

namespace xlDuckDb
{
    internal static class ArrayHelper
    {
        internal static T[,] AsMultiDimensionalArray<T>(this IList<T[]> arrays)
        {
            if (arrays == null) throw new ArgumentNullException(nameof(arrays));
            if (arrays.Count == 0) throw new ArgumentException("Value cannot be an empty collection.", nameof(arrays));

            var minorLength = arrays[0].Length;
            var ret = new T[arrays.Count, minorLength];
            for (var i = 0; i < arrays.Count; i++)
            {
                var array = arrays[i];
                if (array.Length != minorLength)
                    throw new ArgumentException
                        ("All arrays must be the same length");
                for (var j = 0; j < minorLength; j++) ret[i, j] = array[j];
            }

            return ret;
        }
    }
}
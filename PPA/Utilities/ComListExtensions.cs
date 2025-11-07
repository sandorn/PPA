using System;
using System.Collections.Generic;


namespace PPA.Utilities
{
    public static class ComListExtensions
    {
        // 移除了 new() 约束
        public static void DisposeAll<T>(this IEnumerable<T> list) where T : IDisposable
        {
            if (list == null) return;
            foreach (var item in list)
            {
                item?.Dispose();
            }
        }
    }
}

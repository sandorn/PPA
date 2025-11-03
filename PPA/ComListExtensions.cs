using System;
using System.Collections.Generic;


namespace ComListExtensions
{
	public static class ComListExtensions
	{
		// 移除了 new() 约束
		public static void DisposeAll<T>(this IEnumerable<T> list) where T : IDisposable
		{
			if(list==null) return;
			foreach(var item in list)
			{
				// 如果 item 不为 null，就调用其 Dispose 方法
				item?.Dispose();
			}
		}
	}
}

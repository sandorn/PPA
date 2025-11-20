using System;
using System.Collections.Generic;

namespace PPA.Utilities
{
	public static class ComListExtensions
	{
		/// <summary>
		/// 释放所有可释放对象
		/// </summary>
		/// <typeparam name="T"> 可释放对象类型 </typeparam>
		/// <param name="list"> 可释放对象列表 </param>
		public static void DisposeAll<T>(this IEnumerable<T> list) where T : IDisposable
		{
			if(list==null) return;
			foreach(var item in list)
			{
				item?.Dispose();
			}
		}
	}
}

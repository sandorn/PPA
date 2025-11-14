using System;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 图表适配器
	/// </summary>
	public sealed class PowerPointChart : IChart, IComWrapper<NETOP.Chart>
	{
		public IShape ParentShape { get; }
		public NETOP.Chart NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public int ChartType => SafeGet(() => (int) NativeObject.ChartType, 0);

		public PowerPointChart(IShape parent, NETOP.Chart chart)
		{
			ParentShape = parent ?? throw new ArgumentNullException(nameof(parent));
			NativeObject = chart ?? throw new ArgumentNullException(nameof(chart));
		}

		public void ApplyPredefinedStyle(string styleId)
		{
			// 预留：根据 styleId 应用图表样式
		}

		private static T SafeGet<T>(Func<T> getter, T fallback)
		{
			try { return getter(); } catch { return fallback; }
		}
	}
}



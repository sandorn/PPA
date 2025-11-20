using PPA.Core.Abstraction.Presentation;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 图表适配器
	/// </summary>
	public sealed class PowerPointChart(IShape parent,NETOP.Chart chart):IChart, IComWrapper<NETOP.Chart>
	{
		public IShape ParentShape { get; } = parent??throw new ArgumentNullException(nameof(parent));
		public NETOP.Chart NativeObject { get; } = chart??throw new ArgumentNullException(nameof(chart));
		object IComWrapper.NativeObject => NativeObject;

		public int ChartType => ExHandler.SafeGet(() => (int) NativeObject.ChartType,0);

		public void ApplyPredefinedStyle(string styleId)
		{
			// 预留：根据 styleId 应用图表样式
		}
	}
}

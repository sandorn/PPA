using System;
using NetOffice.OfficeApi.Enums;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 形状适配器
	/// </summary>
	public sealed class PowerPointShape : IShape, IComWrapper<NETOP.Shape>
	{
		public IApplication Application { get; }
		public ISlide Slide { get; }
		public NETOP.Shape NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		private readonly IPresentation _presentation;

		public PowerPointShape(IApplication application, IPresentation presentation, ISlide slide, NETOP.Shape shape)
		{
			Application = application ?? throw new ArgumentNullException(nameof(application));
			_presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
			Slide = slide ?? throw new ArgumentNullException(nameof(slide));
			NativeObject = shape ?? throw new ArgumentNullException(nameof(shape));
		}

		public string Name => SafeGet(() => NativeObject?.Name, string.Empty);
		public int ShapeType => SafeGet(() => (int) (NativeObject?.Type ?? 0), 0);
		public bool HasText => SafeGet(() => NativeObject?.TextFrame?.HasText == MsoTriState.msoTrue, false);
		public bool HasTable => SafeGet(() => NativeObject?.HasTable == MsoTriState.msoTrue, false);
		public bool HasChart => SafeGet(() => NativeObject?.HasChart == MsoTriState.msoTrue, false);

		public ITextRange GetTextRange()
		{
			try
			{
				var tr = NativeObject?.TextFrame?.TextRange;
				return tr != null ? new PowerPointTextRange(this,tr) : null;
			} catch { return null; }
		}

		public ITable GetTable()
		{
			try
			{
				var tbl = NativeObject?.Table;
				return tbl != null ? new PowerPointTable(this,tbl) : null;
			} catch { return null; }
		}

		public IChart GetChart()
		{
			try
			{
				var chart = NativeObject?.Chart;
				return chart != null ? new PowerPointChart(this,chart) : null;
			} catch { return null; }
		}

		private static T SafeGet<T>(Func<T> getter, T fallback)
		{
			try { return getter(); } catch { return fallback; }
		}
	}
}



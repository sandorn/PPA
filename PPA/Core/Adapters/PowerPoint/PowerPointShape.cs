using NetOffice.OfficeApi.Enums;
using PPA.Core.Abstraction.Presentation;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 形状适配器
	/// </summary>
	public sealed class PowerPointShape(IApplication application,IPresentation presentation,ISlide slide,NETOP.Shape shape):IShape, IComWrapper<NETOP.Shape>
	{
		public IApplication Application { get; } = application??throw new ArgumentNullException(nameof(application));
		public ISlide Slide { get; } = slide??throw new ArgumentNullException(nameof(slide));
		public NETOP.Shape NativeObject { get; } = shape??throw new ArgumentNullException(nameof(shape));
		object IComWrapper.NativeObject => NativeObject;

		private readonly IPresentation _presentation = presentation??throw new ArgumentNullException(nameof(presentation));

		public string Name => ExHandler.SafeGet(() => NativeObject?.Name,string.Empty);
		public int ShapeType => ExHandler.SafeGet(() => (int) (NativeObject?.Type??0),0);
		public bool HasText => ExHandler.SafeGet(() => NativeObject?.TextFrame?.HasText==MsoTriState.msoTrue,false);
		public bool HasTable => ExHandler.SafeGet(() => NativeObject?.HasTable==MsoTriState.msoTrue,false);
		public bool HasChart => ExHandler.SafeGet(() => NativeObject?.HasChart==MsoTriState.msoTrue,false);

		public ITextRange GetTextRange()
		{
			try
			{
				var tr = NativeObject?.TextFrame?.TextRange;
				return tr!=null ? new PowerPointTextRange(this,tr) : null;
			} catch { return null; }
		}

		public ITable GetTable()
		{
			try
			{
				var tbl = NativeObject?.Table;
				return tbl!=null ? new PowerPointTable(this,tbl) : null;
			} catch { return null; }
		}

		public IChart GetChart()
		{
			try
			{
				var chart = NativeObject?.Chart;
				return chart!=null ? new PowerPointChart(this,chart) : null;
			} catch { return null; }
		}
	}
}

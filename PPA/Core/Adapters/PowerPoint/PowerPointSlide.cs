using System;
using System.Collections.Generic;
using NetOffice.OfficeApi.Enums;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 幻灯片适配器
	/// </summary>
	public sealed class PowerPointSlide : ISlide, IComWrapper<NETOP.Slide>
	{
		public IApplication Application { get; }
		public IPresentation Presentation { get; }
		public NETOP.Slide NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public string Title
		{
			get
			{
				try
				{
					var titleShape = NativeObject.Shapes?.Title;
					if(titleShape!=null && titleShape.TextFrame?.HasText==MsoTriState.msoTrue)
					{
						return titleShape.TextFrame.TextRange?.Text ?? string.Empty;
					}
				} catch { }
				return string.Empty;
			}
		}

		public int SlideIndex => SafeGet(() => NativeObject?.SlideIndex ?? 0, 0);

		public PowerPointSlide(IApplication application, IPresentation presentation, NETOP.Slide slide)
		{
			Application = application ?? throw new ArgumentNullException(nameof(application));
			Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
			NativeObject = slide ?? throw new ArgumentNullException(nameof(slide));
		}

		public IReadOnlyList<IShape> Shapes
		{
			get
			{
				var list = new List<IShape>();
				try
				{
					foreach(NETOP.Shape s in NativeObject.Shapes)
					{
						list.Add(new PowerPointShape(Application, Presentation, this, s));
					}
				} catch { /* ignore */ }
				return list;
			}
		}

		public IShape FindShapeByName(string name)
		{
			if(string.IsNullOrWhiteSpace(name)) return null;
			try
			{
				var shape = NativeObject.Shapes[name];
				return shape != null ? new PowerPointShape(Application, Presentation, this, shape) : null;
			} catch
			{
				return null;
			}
		}

		private static T SafeGet<T>(Func<T> getter, T fallback)
		{
			try { return getter(); } catch { return fallback; }
		}
	}
}



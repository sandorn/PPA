using NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 演示文稿适配器
	/// </summary>
	public sealed class PowerPointPresentation(IApplication application,NETOP.Presentation presentation):IPresentation, IComWrapper<NETOP.Presentation>
	{
		public IApplication Application { get; } = application??throw new ArgumentNullException(nameof(application));
		public NETOP.Presentation NativeObject { get; } = presentation??throw new ArgumentNullException(nameof(presentation));
		object IComWrapper.NativeObject => NativeObject;

		public string Name => ExHandler.SafeGet(() => NativeObject?.Name,string.Empty);
		public int SlideCount => ExHandler.SafeGet(() => NativeObject?.Slides?.Count??0,0);

		public IReadOnlyList<ISlide> Slides
		{
			get
			{
				var list = new List<ISlide>();
				try
				{
					foreach(NETOP.Slide s in NativeObject.Slides.Cast<Slide>())
					{
						list.Add(new PowerPointSlide(Application,this,s));
					}
				} catch { /* ignore */ }
				return list;
			}
		}

		public ISlide GetSlide(int index)
		{
			try
			{
				var s = NativeObject.Slides[index];
				return s!=null ? new PowerPointSlide(Application,this,s) : null;
			} catch
			{
				return null;
			}
		}
	}
}

using System;
using System.Collections.Generic;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 演示文稿适配器
	/// </summary>
	public sealed class PowerPointPresentation : IPresentation, IComWrapper<NETOP.Presentation>
	{
		public IApplication Application { get; }
		public NETOP.Presentation NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public string Name => SafeGet(() => NativeObject?.Name, string.Empty);
		public int SlideCount => SafeGet(() => NativeObject?.Slides?.Count ?? 0, 0);

		public PowerPointPresentation(IApplication application, NETOP.Presentation presentation)
		{
			Application = application ?? throw new ArgumentNullException(nameof(application));
			NativeObject = presentation ?? throw new ArgumentNullException(nameof(presentation));
		}

		public IReadOnlyList<ISlide> Slides
		{
			get
			{
				var list = new List<ISlide>();
				try
				{
					foreach(NETOP.Slide s in NativeObject.Slides)
					{
						list.Add(new PowerPointSlide(Application, this, s));
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
				return s != null ? new PowerPointSlide(Application, this, s) : null;
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



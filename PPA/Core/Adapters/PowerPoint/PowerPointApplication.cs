using PPA.Core.Abstraction.Presentation;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 应用程序适配器
	/// </summary>
	public sealed class PowerPointApplication:IApplication, IComWrapper<NETOP.Application>
	{
		public NETOP.Application NativeObject { get; }
		object IComWrapper.NativeObject => NativeObject;

		public ApplicationType ApplicationType => ApplicationType.PowerPoint;

		public PowerPointApplication(NETOP.Application app)
		{
			NativeObject=app??throw new ArgumentNullException(nameof(app));
		}

		public IPresentation GetActivePresentation()
		{
			try
			{
				var p = NativeObject?.ActivePresentation;
				return p!=null ? new PowerPointPresentation(this,p) : null;
			} catch
			{
				return null;
			}
		}

		public ISlide GetActiveSlide()
		{
			try
			{
				var view = NativeObject?.ActiveWindow?.View;
				var slide = view?.Slide as NETOP.Slide;
				return slide!=null ? new PowerPointSlide(this,GetActivePresentationInternal(slide),slide) : null;
			} catch
			{
				return null;
			}
		}

		public IReadOnlyList<IShape> GetSelectedShapes()
		{
			try
			{
				var selection = NativeObject?.ActiveWindow?.Selection;
				if(selection==null) return Array.Empty<IShape>();

				// 选中单个形状
				if(selection.Type==NETOP.Enums.PpSelectionType.ppSelectionShapes)
				{
					var range = selection.ShapeRange;
					if(range==null) return Array.Empty<IShape>();
					var presentation = GetActivePresentation();
					var slide = GetActiveSlide();
					var list = new List<IShape>();
					foreach(NETOP.Shape s in range)
					{
						list.Add(new PowerPointShape(this,presentation,slide,s));
					}
					return list;
				}

				// 选中文本
				if(selection.Type==NETOP.Enums.PpSelectionType.ppSelectionText)
				{
					var shape = selection.TextRange?.Parent as NETOP.Shape;
					if(shape!=null)
					{
						var presentation = GetActivePresentation();
						var slide = GetActiveSlide();
						return new[] { new PowerPointShape(this,presentation,slide,shape) };
					}
				}

				return Array.Empty<IShape>();
			} catch
			{
				return Array.Empty<IShape>();
			}
		}

		public FeatureSupportLevel GetFeatureSupport(string featureKey)
		{
			// 先统一标 Full；后续按特性细化
			return FeatureSupportLevel.Full;
		}

		public void RunOnUiThread(Action action)
		{
			// VSTO 单线程模型：直接执行
			action?.Invoke();
		}

		private IPresentation GetActivePresentationInternal(NETOP.Slide slide)
		{
			try
			{
				return slide?.Parent is NETOP.Presentation p ? new PowerPointPresentation(this,p) : null;
			} catch
			{
				return null;
			}
		}
	}
}

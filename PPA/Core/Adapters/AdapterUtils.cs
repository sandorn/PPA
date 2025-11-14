using System;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters
{
	/// <summary>
	/// 统一的抽象装配工具，减少各处重复的包装代码
	/// </summary>
	public static class AdapterUtils
	{
		public static IApplication WrapApplication(NETOP.Application app)
		{
			if(app==null) return null;
			return new PowerPointApplication(app);
		}

		public static ISlide WrapSlide(NETOP.Application app, NETOP.Shape shape)
		{
			if(app==null||shape==null) return null;
			var iApp = WrapApplication(app);
			NETOP.Slide nativeSlide = null;
			try { nativeSlide = shape.Parent as NETOP.Slide; } catch { /* ignore */ }
			if(nativeSlide==null) return null;
			var iPres = WrapPresentation(iApp,nativeSlide);
			return new PowerPointSlide(iApp,iPres,nativeSlide);
		}

		public static ISlide WrapSlide(NETOP.Application app, NETOP.Slide slide)
		{
			if(app==null||slide==null) return null;
			var iApp = WrapApplication(app);
			var iPres = WrapPresentation(iApp,slide);
			return new PowerPointSlide(iApp,iPres,slide);
		}

		public static IPresentation WrapPresentation(IApplication iApp, NETOP.Slide nativeSlide)
		{
			if(iApp==null||nativeSlide==null) return null;
			NETOP.Presentation nativePres = null;
			try { nativePres = nativeSlide.Parent as NETOP.Presentation; } catch { /* ignore */ }
			return nativePres!=null ? new PowerPointPresentation(iApp,nativePres) : null;
		}


		public static IShape WrapShape(NETOP.Application app, NETOP.Shape shape)
		{
			PPA.Core.Profiler.LogMessage($"WrapShape(NETOP.Application) 被调用，app类型={app?.GetType().Name ?? "null"}, shape类型={shape?.GetType().Name ?? "null"}", "ADAPTER");
			if(app==null||shape==null)
			{
				PPA.Core.Profiler.LogMessage("WrapShape: app 或 shape 为 null，返回 null", "ADAPTER_WARN");
				return null;
			}

			PPA.Core.Profiler.LogMessage("使用 PowerPoint 适配器包装形状", "ADAPTER");
			var iApp = WrapApplication(app);
			NETOP.Slide nativeSlide = null;
			try { nativeSlide = shape.Parent as NETOP.Slide; } catch { /* ignore */ }
			var iPres = WrapPresentation(iApp,nativeSlide);
			var iSlide = nativeSlide!=null ? new PowerPointSlide(iApp,iPres,nativeSlide) : null;
			return new PowerPointShape(iApp,iPres,iSlide,shape);
		}

		public static ITable WrapTable(NETOP.Application app, NETOP.Shape shape, NETOP.Table table)
		{
			PPA.Core.Profiler.LogMessage($"WrapTable 被调用，app类型={app?.GetType().Name ?? "null"}, shape类型={shape?.GetType().Name ?? "null"}, table类型={table?.GetType().Name ?? "null"}", "ADAPTER");
			if(app==null||shape==null||table==null)
			{
				PPA.Core.Profiler.LogMessage("WrapTable: app、shape 或 table 为 null，返回 null", "ADAPTER_WARN");
				return null;
			}

			PPA.Core.Profiler.LogMessage("使用 PowerPoint 适配器包装表格", "ADAPTER");
			var iShape = WrapShape(app,shape);
			if(iShape == null)
			{
				PPA.Core.Profiler.LogMessage("WrapShape 返回 null，无法创建表格", "ADAPTER_ERROR");
				return null;
			}
			return new PowerPointTable(iShape,table);
		}

		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}
	}
}



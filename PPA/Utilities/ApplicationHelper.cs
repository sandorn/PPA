using PPA.Core;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using NETOP = NetOffice.PowerPointApi;
using MSOP = Microsoft.Office.Interop.PowerPoint;

namespace PPA.Utilities
{
	/// <summary>
	/// PowerPoint Application 对象获取辅助类
	/// 提供统一的 Application 对象获取方法，避免代码重复
	/// </summary>
	public static class ApplicationHelper
	{
		/// <summary>
		/// 获取 NetOffice PowerPoint 应用程序对象
		/// </summary>
		/// <remarks>
		/// 此方法返回的是 NetOffice 包装的 Application 对象（NETOP.Application），
		/// 而不是原生 COM 对象。NetOffice 提供了更友好的 API 和更好的异常处理。
		/// 如果需要原生 COM 对象（MSOP.Application），请使用 GetNativeComApplication() 方法。
		/// </remarks>
		/// <param name="application">可选的 IApplication 接口，如果提供则优先使用</param>
		/// <returns>NetOffice Application 对象，如果无法获取则返回 null</returns>
		public static NETOP.Application GetNetOfficeApplication(IApplication application = null)
		{
			try
			{
				// 如果提供了 IApplication 接口，尝试从中获取 NetOffice 对象
				if(application != null)
				{
					if(application is IComWrapper<NETOP.Application> typed)
					{
						return typed.NativeObject;
					}

					if(application is IComWrapper wrapper)
					{
						return wrapper.NativeObject as NETOP.Application;
					}
				}

				// 尝试从全局 ThisAddIn 获取
				var addIn = Globals.ThisAddIn;
				if(addIn?.NetApp != null)
				{
					return addIn.NetApp;
				}

				// 尝试从 DI 容器获取
				var serviceProvider = addIn?.ServiceProvider;
				if(serviceProvider != null)
				{
					try
					{
						var factoryObj = serviceProvider.GetService(typeof(IApplicationFactory)) as IApplicationFactory;
						var resolvedFromFactory = factoryObj?.GetCurrent() as IComWrapper<NETOP.Application>;
						if(resolvedFromFactory != null)
						{
							return resolvedFromFactory.NativeObject;
						}
					}
					catch
					{
						// 忽略异常，继续尝试其他方式
					}
				}

				return null;
			}
			catch(System.Exception ex)
			{
				Profiler.LogMessage($"ApplicationHelper.GetNetOfficeApplication: 获取 NetOffice Application 对象失败: {ex.Message}", "ERROR");
				return null;
			}
		}

		/// <summary>
		/// 获取原生 COM Application 对象
		/// </summary>
		/// <remarks>
		/// 此方法返回的是原生 COM 对象（MSOP.Application，即 Microsoft.Office.Interop.PowerPoint.Application），
		/// 而不是 NetOffice 包装的对象。原生 COM 对象在某些场景下需要直接访问底层 COM 接口。
		/// 如果需要 NetOffice 对象（NETOP.Application），请使用 GetNetOfficeApplication() 方法。
		/// </remarks>
		/// <param name="application">可选的 IApplication 接口，如果提供则优先使用</param>
		/// <returns>原生 COM Application 对象，如果无法获取则返回 null</returns>
		public static MSOP.Application GetNativeComApplication(IApplication application = null)
		{
			try
			{
				// 优先从全局 ThisAddIn 获取原生 COM 对象
				var addIn = Globals.ThisAddIn;
				if(addIn?.NativeApp != null)
				{
					return addIn.NativeApp;
				}

				// 从 NetOffice 对象获取底层原生 COM 对象
				var netApp = GetNetOfficeApplication(application);
				if(netApp != null)
				{
					return netApp.UnderlyingObject as MSOP.Application;
				}

				return null;
			}
			catch(System.Exception ex)
			{
				Profiler.LogMessage($"ApplicationHelper.GetNativeComApplication: 获取原生 COM Application 对象失败: {ex.Message}", "ERROR");
				return null;
			}
		}
	}
}


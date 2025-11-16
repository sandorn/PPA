using PPA.Core;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Adapters.PowerPoint;
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

		/// <summary>
		/// 从 NetOffice Application 对象获取原生 COM Application 对象
		/// </summary>
		/// <remarks>
		/// 此重载方法直接从 NETOP.Application 对象获取其底层的原生 COM 对象。
		/// 用于需要直接访问底层 COM 接口的场景，例如避免 NetOffice 包装本地化类名的问题。
		/// </remarks>
		/// <param name="netApp">NetOffice Application 对象</param>
		/// <returns>原生 COM Application 对象，如果 netApp 为 null 或无法获取则返回 null</returns>
		public static MSOP.Application GetNativeComApplication(NETOP.Application netApp)
		{
			if(netApp == null) return null;

			try
			{
				// 优先从全局 ThisAddIn 获取原生 COM 对象（如果匹配）
				var addIn = Globals.ThisAddIn;
				if(addIn?.NetApp == netApp && addIn?.NativeApp != null)
				{
					return addIn.NativeApp;
				}

				// 从 NetOffice 对象获取底层原生 COM 对象
				return (netApp as NetOffice.ICOMObject)?.UnderlyingObject as MSOP.Application;
			}
			catch(System.Exception ex)
			{
				Profiler.LogMessage($"ApplicationHelper.GetNativeComApplication: 从 NetOffice Application 获取原生 COM 对象失败: {ex.Message}", "ERROR");
				return null;
			}
		}

		/// <summary>
		/// 将 NetOffice Application 对象转换为抽象接口
		/// </summary>
		/// <remarks>
		/// 此方法将 NETOP.Application 包装为 IApplication 接口。
		/// 如果 netApp 已经是某个 IApplication 适配器的底层对象，会尝试复用该适配器。
		/// 否则会创建新的 PowerPointApplication 适配器。
		/// </remarks>
		/// <param name="netApp">NetOffice Application 对象</param>
		/// <returns>IApplication 接口对象，如果 netApp 为 null 则返回 null</returns>
		public static IApplication GetAbstractApplication(NETOP.Application netApp)
		{
			if(netApp == null) return null;

			// 尝试从 DI 容器查找已存在的适配器
			var addIn = Globals.ThisAddIn;
			var serviceProvider = addIn?.ServiceProvider;
			if(serviceProvider != null)
			{
				try
				{
					var factoryObj = serviceProvider.GetService(typeof(IApplicationFactory)) as IApplicationFactory;
					var existingApp = factoryObj?.GetCurrent();
					if(existingApp is IComWrapper<NETOP.Application> typed && typed.NativeObject == netApp)
					{
						return existingApp;
					}
				}
				catch
				{
					// 忽略异常，继续创建新适配器
				}
			}

			// 创建新的适配器对象
			return new PowerPointApplication(netApp);
		}
	}
}


using PPA.Core;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Adapters.PowerPoint;
using PPA.Core.Logging;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Utilities
{
	/// <summary>
	/// PowerPoint Application 对象获取辅助类 提供统一的 Application 对象获取方法，避免代码重复
	/// </summary>
	public static class ApplicationHelper
	{
		private static readonly ILogger _logger = LoggerProvider.GetLogger();

		/// <summary>
		/// 获取 NetOffice PowerPoint 应用程序对象
		/// </summary>
		/// <remarks>
		/// 此方法返回的是 NetOffice 包装的 Application 对象（NETOP.Application）， 而不是原生 COM 对象。NetOffice 提供了更友好的
		/// API 和更好的异常处理。 如果需要原生 COM 对象（MSOP.Application），请使用 GetNativeComApplication() 方法。
		/// </remarks>
		/// <param name="application"> 可选的 IApplication 接口，如果提供则优先使用 </param>
		/// <returns> NetOffice Application 对象，如果无法获取则返回 null </returns>
		public static NETOP.Application GetNetOfficeApplication(IApplication application = null)
		{
			return NetChannel.Resolve(application);
		}

		/// <summary>
		/// 获取原生 COM Application 对象
		/// </summary>
		/// <remarks>
		/// 此方法返回的是原生 COM 对象（MSOP.Application，即 Microsoft.Office.Interop.PowerPoint.Application），
		/// 而不是 NetOffice 包装的对象。原生 COM 对象在某些场景下需要直接访问底层 COM 接口。 如果需要 NetOffice
		/// 对象（NETOP.Application），请使用 GetNetOfficeApplication() 方法。
		/// </remarks>
		/// <param name="application"> 可选的 IApplication 接口，如果提供则优先使用 </param>
		/// <returns> 原生 COM Application 对象，如果无法获取则返回 null </returns>
		public static MSOP.Application GetNativeComApplication(IApplication application = null)
		{
			return NativeChannel.Resolve(application);
		}

		/// <summary>
		/// 从 NetOffice Application 对象获取原生 COM Application 对象
		/// </summary>
		/// <remarks>
		/// 此重载方法直接从 NETOP.Application 对象获取其底层的原生 COM 对象。 用于需要直接访问底层 COM 接口的场景，例如避免 NetOffice 包装本地化类名的问题。
		/// </remarks>
		/// <param name="netApp"> NetOffice Application 对象 </param>
		/// <returns> 原生 COM Application 对象，如果 netApp 为 null 或无法获取则返回 null </returns>
		public static MSOP.Application GetNativeComApplication(NETOP.Application netApp)
		{
			return NativeChannel.Resolve(netApp);
		}

		/// <summary>
		/// 将 NetOffice Application 对象转换为抽象接口
		/// </summary>
		/// <remarks>
		/// 此方法将 NETOP.Application 包装为 IApplication 接口。 如果 netApp 已经是某个 IApplication
		/// 适配器的底层对象，会尝试复用该适配器。 否则会创建新的 PowerPointApplication 适配器。
		/// </remarks>
		/// <param name="netApp"> NetOffice Application 对象 </param>
		/// <returns> IApplication 接口对象，如果 netApp 为 null 则返回 null </returns>
		public static IApplication GetAbstractApplication(NETOP.Application netApp)
		{
			if(netApp==null) return null;

			// 尝试从 DI 容器查找已存在的适配器
			var provider = ApplicationProvider.Current;
			var serviceProvider = provider?.ServiceProvider;
			if(serviceProvider!=null)
			{
				try
				{
					var factoryObj = serviceProvider.GetService(typeof(IApplicationFactory)) as IApplicationFactory;
					var existingApp = factoryObj?.GetCurrent();
					if(existingApp!=null&&AdapterUtils.UnwrapApplication(existingApp)==netApp)
					{
						return existingApp;
					}
				} catch
				{
					// 忽略异常，继续创建新适配器
				}
			}

			// 创建新的适配器对象
			return new PowerPointApplication(netApp);
		}

		/// <summary>
		/// 确保返回一个可用的 NetOffice Application（ActiveWindow 失效时自动刷新）
		/// </summary>
		/// <param name="netApp"> 现有的 NetOffice Application 对象 </param>
		/// <returns> 可用的 NetOffice Application；若无法获取则返回 null </returns>
		public static NETOP.Application EnsureValidNetApplication(NETOP.Application netApp)
		{
			return NetChannel.EnsureValid(netApp);
		}

		/// <summary>
		/// 从 IApplication 接口获取并确保返回一个可用的 NetOffice Application
		/// </summary>
		/// <param name="application"> IApplication 接口对象 </param>
		/// <returns> 可用的 NetOffice Application；若无法获取则返回 null </returns>
		public static NETOP.Application EnsureValidNetApplication(IApplication application)
		{
			var netApp = GetNetOfficeApplication(application);
			return EnsureValidNetApplication(netApp);
		}

		private static class NetChannel
		{
			public static NETOP.Application Resolve(IApplication application)
			{
				try
				{
					if(application!=null)
					{
						var unwrapped = AdapterUtils.UnwrapApplication(application);
						if(unwrapped!=null)
						{
							return unwrapped;
						}
					}

					var provider = ApplicationProvider.Current;
					if(provider?.NetApplication!=null)
					{
						return provider.NetApplication;
					}

					var serviceProvider = provider?.ServiceProvider;
					if(serviceProvider!=null)
					{
						var factoryObj = serviceProvider.GetService(typeof(IApplicationFactory)) as IApplicationFactory;
						var resolvedFromFactory = factoryObj?.GetCurrent();
						if(resolvedFromFactory!=null)
						{
							return AdapterUtils.UnwrapApplication(resolvedFromFactory);
						}
					}

					return null;
				} catch(System.Exception ex)
				{
					_logger.LogError($"获取 NetOffice Application 对象失败: {ex.Message}",ex);
					return null;
				}
			}

			public static NETOP.Application EnsureValid(NETOP.Application netApp)
			{
				if(netApp==null)
				{
					var fallback = Resolve(null);
					if(fallback==null)
					{
						_logger.LogWarning("Application 为 null 且无法重新获取");
					}
					return fallback;
				}

				var window = ExHandler.SafeGet(() => netApp.ActiveWindow, defaultValue:(NETOP.DocumentWindow)null);
				if(window!=null)
				{
					return netApp;
				}

				var refreshed = Resolve(null);
				if(refreshed!=null)
				{
					_logger.LogInformation("ActiveWindow 无效，已重新获取 Application");
					return refreshed;
				}

				_logger.LogWarning("ActiveWindow 无效且无法重新获取 Application");
				return null;
			}
		}

		private static class NativeChannel
		{
			public static MSOP.Application Resolve(IApplication application)
			{
				try
				{
					var provider = ApplicationProvider.Current;
					if(provider?.NativeApplication!=null)
					{
						return provider.NativeApplication;
					}

					var netApp = NetChannel.Resolve(application);
					return Resolve(netApp);
				} catch(System.Exception ex)
				{
					_logger.LogError($"获取原生 COM Application 对象失败: {ex.Message}",ex);
					return null;
				}
			}

			public static MSOP.Application Resolve(NETOP.Application netApp)
			{
				if(netApp==null)
				{
					_logger.LogWarning("netApp 为 null");
					return null;
				}

				try
				{
					var provider = ApplicationProvider.Current;
					if(provider?.NativeApplication!=null&&MatchesProviderNetApp(provider,netApp))
					{
						_logger.LogDebug("从 ApplicationProvider 获取到 NativeApplication");
						return provider.NativeApplication;
					}

					return ExtractUnderlying(netApp);
				} catch(System.Exception ex)
				{
					_logger.LogError($"从 NetOffice Application 获取原生 COM 对象失败: {ex.Message}",ex);
					return null;
				}
			}

			private static bool MatchesProviderNetApp(IApplicationProvider provider,NETOP.Application netApp)
			{
				return ExHandler.SafeGet(() =>
				{
					var _ = netApp.Name;
					if(provider.NetApplication!=null)
					{
						var __ = provider.NetApplication.Name;
						return ReferenceEquals(provider.NetApplication,netApp);
					}
					return false;
				},defaultValue: false);
			}

			private static MSOP.Application ExtractUnderlying(NETOP.Application netApp)
			{
				var result = ExHandler.SafeGet(() =>
				{
					var comObject = netApp as NetOffice.ICOMObject;
					var underlying = comObject?.UnderlyingObject as MSOP.Application;
					if(underlying != null)
					{
						_logger.LogDebug("从 UnderlyingObject 获取到原生 COM 对象");
					}
					return underlying;
				}, defaultValue: (MSOP.Application)null);

				if(result==null)
				{
					_logger.LogWarning("所有方法都失败，返回 null");
				}

				return result;
			}
		}
	}
}

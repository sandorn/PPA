using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.DI;
using PPA.Core.Logging;
using PPA.Manipulation;
using PPA.UI;
using System;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA
{
	/// <summary>
	/// PowerPoint 插件的主入口类 处理插件的初始化、资源管理和事件响应
	/// </summary>
	public partial class ThisAddIn
	{
		#region Private Fields

		private CustomRibbon _customRibbon; // 自定义功能区实例
		private ApplicationProvider _applicationProvider; // 应用程序上下文提供者
		private bool _resourcesDisposed = false; // 资源是否已释放的标记
		public MSOP.Application NativeApp { get; private set; } // 本地PowerPoint应用程序实例
		public NETOP.Application NetApp { get; private set; } // NetOffice PowerPoint 应用程序实例
		private IServiceProvider _serviceProvider; // DI 容器服务提供者
		private readonly ILogger _fallbackLogger = new ProfilerLoggerAdapter(); // 基础日志
		private ILogger _logger;

		private ILogger Logger => _logger??_fallbackLogger;

		/// <summary>
		/// 获取 DI 容器服务提供者（用于向后兼容的静态方法）
		/// </summary>
		internal IServiceProvider ServiceProvider => _serviceProvider;

		#endregion Private Fields

		#region Protected Methods

		/// <summary>
		/// 创建功能区扩展性对象
		/// </summary>
		/// <returns> 自定义功能区实例 </returns>
		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			Logger.LogInformation("创建Ribbon对象");

			// 此时 App 可能还没有初始化，所以传递 null
			_customRibbon=new CustomRibbon();
			return _customRibbon;
		}

		#endregion Protected Methods

		#region Private Methods

		/// <summary>
		/// 插件关闭时的事件处理程序
		/// </summary>
		/// <param name="sender"> 事件发送者 </param>
		/// <param name="e"> 事件参数 </param>
		private void ThisAddIn_Shutdown(object sender,System.EventArgs e)
		{
			Logger.LogInformation("插件正在关闭");

			// 注销快捷键系统
			KeyboardShortcutHelper.Uninitialize();

			CleanupResources();
		}

		/// <summary>
		/// 清理插件使用的所有资源 确保正确释放COM对象，避免内存泄漏
		/// </summary>
		private void CleanupResources()
		{
			if(_resourcesDisposed) return; // 防止重复清理

			try
			{
				// 释放功能区资源
				_customRibbon?.Dispose();
				_customRibbon=null;

				// 释放 DI 容器
				if(_serviceProvider is IDisposable disposableServiceProvider)
				{
					try
					{
						disposableServiceProvider.Dispose();
					} catch(Exception ex)
					{
						Logger.LogWarning($"释放 DI 容器时出错: {ex.Message}");
					} finally
					{
						_serviceProvider=null;
					}
				}

				// 释放NetOffice应用程序实例
				if(NetApp!=null)
				{
					try
					{
						NetApp.Dispose();
					} catch(Exception ex)
					{
						Logger.LogWarning($"释放App对象时出错: {ex.Message}");
					} finally
					{
						NetApp=null;
					}
				}
			} finally
			{
				_resourcesDisposed=true;
				UpdateApplicationProviderContext();
			}
		}

		/// <summary>
		/// 插件启动时的事件处理程序
		/// </summary>
		/// <param name="sender"> 事件发送者 </param>
		/// <param name="e"> 事件参数 </param>
		private void ThisAddIn_Startup(object sender,System.EventArgs e)
		{
			Logger.LogInformation("插件正在启动");

			// 初始化多语言资源管理器
			ResourceManager.Initialize("PPA.Properties.Resources",System.Reflection.Assembly.GetExecutingAssembly());

			// 初始化 DI 容器（在 Application 初始化之前，先注册基础服务）
			InitializeDIContainer();

			// 测试 DI 容器（可选，用于验证）
			TestDIContainer();

			InitializeNetOfficeApplication();
			UpdateApplicationProviderContext();

			// Startup 完成后，将 App 设置到 CustomRibbon
			_customRibbon?.SetApplicationProvider(_applicationProvider);
			_customRibbon?.SetApplication(NetApp);

			// 初始化快捷键系统
			KeyboardShortcutHelper.Initialize(_applicationProvider);
		}

		/// <summary>
		/// 初始化 DI 容器 注册所有 PPA 服务到依赖注入容器
		/// </summary>
		private void InitializeDIContainer()
		{
			try
			{
				// 启用文件日志
				Profiler.EnableFileLogging=true;
				var logPath = System.IO.Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
					"PPA",
					$"PPA_{DateTime.Now:yyyyMMdd_HHmmss}.log"
				);
				// 确保目录存在
				var logDir = System.IO.Path.GetDirectoryName(logPath);
				if(!System.IO.Directory.Exists(logDir))
				{
					System.IO.Directory.CreateDirectory(logDir);
				}
				Profiler.LogFilePath=logPath;
				Logger.LogInformation($"日志文件路径: {logPath}");
				Profiler.CleanupLogFiles(logDir);
				
				var services = new ServiceCollection();
				services.AddPPAServices();
				_serviceProvider=services.BuildServiceProvider();
				_logger=_serviceProvider.GetService<ILogger>()??_fallbackLogger;
				_applicationProvider=_serviceProvider.GetService<ApplicationProvider>();
				
				// 应用配置中的日志策略
				try
				{
					var formattingConfig = _serviceProvider.GetService<PPA.Core.Abstraction.Business.IFormattingConfig>() as FormattingConfig;
					formattingConfig?.ApplyLoggingConfigToProfiler();
				}
				catch(Exception ex)
				{
					Logger.LogWarning($"应用日志配置失败: {ex.Message}");
				}
				UpdateApplicationProviderContext();
				Logger.LogInformation("DI 容器初始化成功");
			} catch(Exception ex)
			{
				Logger.LogError($"初始化 DI 容器失败: {ex.Message}",ex);
				Logger.LogDebug($"堆栈跟踪: {ex.StackTrace}");
			}
		}

		/// <summary>
		/// 测试 DI 容器是否正常工作
		/// </summary>
		private void TestDIContainer()
		{
			if(_serviceProvider==null)
			{
				Logger.LogWarning("DI 容器未初始化");
				return;
			}

			try
			{
				var config = _serviceProvider.GetService<PPA.Core.Abstraction.Business.IFormattingConfig>();
				if(config!=null)
				{
					Logger.LogInformation("DI 容器测试成功：可以获取 IFormattingConfig 服务");
				} else
				{
					Logger.LogWarning("DI 容器测试失败：无法获取 IFormattingConfig 服务");
				}
			} catch(Exception ex)
			{
				Logger.LogError($"DI 容器测试失败: {ex.Message}",ex);
			}
		}

		/// <summary>
		/// 初始化NetOffice应用程序实例 创建基于本地PowerPoint应用的包装器
		/// </summary>
		private void InitializeNetOfficeApplication()
		{
			try
			{
				// 获取原生 PowerPoint 应用程序对象（Application 属性在 ThisAddIn.Designer.cs 中定义）
				NativeApp=this.Application;

				if(NativeApp==null)
				{
					Logger.LogWarning("本地PowerPoint应用程序对象为空");
					return;
				}

				NetApp=new NETOP.Application(null,NativeApp);

				if(NetApp!=null)
				{
					Logger.LogInformation("NetOffice包装器初始化成功");
				}
			} catch(Exception ex)
			{
				Logger.LogError($"初始化NetOffice应用程序失败: {ex.Message}",ex);
				Logger.LogDebug($"堆栈跟踪: {ex.StackTrace}");
			}
		}

		private void UpdateApplicationProviderContext()
		{
			_applicationProvider?.SetContext(NetApp,NativeApp,_serviceProvider);
		}

		#endregion Private Methods

		#region VSTO Generated Code

		/// <summary>
		/// VSTO自动生成的启动代码 注册启动和关闭事件处理程序
		/// </summary>
		private void InternalStartup()
		{
			Logger.LogInformation("内部启动过程");
			this.Startup+=ThisAddIn_Startup;
			this.Shutdown+=ThisAddIn_Shutdown;
		}

		#endregion VSTO Generated Code
	}
}

using System.Diagnostics;
using Office = Microsoft.Office.Core;
using NETOP = NetOffice.PowerPointApi;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;

namespace PPA
{
	/// <summary>
	/// PowerPoint 插件的主入口类
	/// 处理插件的初始化、资源管理和事件响应
	/// </summary>
	public partial class ThisAddIn
	{
		#region Private Fields

		private CustomRibbon _customRibbon; // 自定义功能区实例
		private bool _resourcesDisposed = false; // 资源是否已释放的标记
		public PowerPoint.Application NativeApp { get; private set; } // 本地PowerPoint应用程序实例
		public NETOP.Application NetApp { get; private set; } // NetOffice PowerPoint 应用程序实例

		#endregion Private Fields

		#region Protected Methods

		/// <summary>
		/// 创建功能区扩展性对象
		/// </summary>
		/// <returns>自定义功能区实例</returns>
		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			Debug.WriteLine("[ThisAddIn] 创建功能区扩展性对象");

			// 此时 App 可能还没有初始化，所以传递 null
			_customRibbon = new CustomRibbon();
			return _customRibbon;
		}

		#endregion Protected Methods

		#region Private Methods

		/// <summary>
		/// 插件关闭时的事件处理程序
		/// </summary>
		/// <param name="sender">事件发送者</param>
		/// <param name="e">事件参数</param>
		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			Debug.WriteLine("[ThisAddIn] 插件正在关闭");
			CleanupResources();
		}

		/// <summary>
		/// 清理插件使用的所有资源
		/// 确保正确释放COM对象，避免内存泄漏
		/// </summary>
		private void CleanupResources()
		{
			if (_resourcesDisposed) return; // 防止重复清理

			try
			{
				// 释放功能区资源
				_customRibbon?.Dispose();
				_customRibbon = null;

				// 释放NetOffice应用程序实例
				if (NetApp != null)
				{
					try
					{
						NetApp.Dispose();
					} catch (Exception ex)
					{
						Debug.WriteLine($"[ThisAddIn] 释放App对象时出错: {ex.Message}");
					} finally
					{
						NetApp = null;
					}
				}
			} finally
			{
				_resourcesDisposed = true;
			}
		}

		/// <summary>
		/// 插件启动时的事件处理程序
		/// </summary>
		/// <param name="sender">事件发送者</param>
		/// <param name="e">事件参数</param>
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			Debug.WriteLine("[ThisAddIn] 插件正在启动");
			InitializeNetOfficeApplication();

			// Startup 完成后，将 App 设置到 CustomRibbon
			_customRibbon?.SetApplication(NetApp);
		}

		/// <summary>
		/// 初始化NetOffice应用程序实例
		/// 创建基于本地PowerPoint应用的包装器
		/// </summary>
		private void InitializeNetOfficeApplication()
		{
			try
			{
				NativeApp= Globals.ThisAddIn.Application;

				if (NativeApp== null)
				{
					Debug.WriteLine("[ThisAddIn] 本地PowerPoint应用程序对象为空");
					return;
				}

				NetApp = new NETOP.Application(null,NativeApp);

				if (NetApp != null)
				{
					Debug.WriteLine("[ThisAddIn] NetOffice包装器初始化成功");
				}
			} catch (Exception ex)
			{
				Debug.WriteLine($"[ThisAddIn] 初始化NetOffice应用程序失败: {ex.Message}");
				Debug.WriteLine($"[ThisAddIn] 堆栈跟踪: {ex.StackTrace}");
			}
		}

		#endregion Private Methods

		#region VSTO Generated Code

		/// <summary>
		/// VSTO自动生成的启动代码
		/// 注册启动和关闭事件处理程序
		/// </summary>
		private void InternalStartup()
		{
			Debug.WriteLine("[ThisAddIn] 内部启动过程");
			this.Startup += ThisAddIn_Startup;
			this.Shutdown += ThisAddIn_Shutdown;
		}

		#endregion VSTO Generated Code
	}
}
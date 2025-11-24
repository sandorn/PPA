using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using PPA.Properties;
using PPA.UI.Forms;
using PPA.UI.Providers;
using PPA.Utilities;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA
{
	[ComVisible(true)]
	public class CustomRibbon:Office.IRibbonExtensibility, IDisposable
	{
		#region Private Fields

		private Office.IRibbonUI _ribbonUI;
		private NETOP.Application _netApp; // NetOffice Application 对象
		private bool _tb101Press;
		private bool _disposed = false;
		private bool _appInitialized = false;
		private IApplicationProvider _applicationProvider;
		private int _lastShapeCount = -1; // 记录上次选中的形状数量，用于检测变化
		private ILogger _logger = LoggerProvider.GetLogger();
		private ILogger Logger => _logger??LoggerProvider.GetLogger();

		// Ribbon 相关服务
		private IRibbonXmlProvider _ribbonXmlProvider;

		private IRibbonIconProvider _ribbonIconProvider;
		private IRibbonCommandRouter _ribbonCommandRouter;
		private ISelectionService _selectionService; // 新增字段

		#endregion Private Fields

		#region Initialization & Setup

		/// <summary>
		/// CustomRibbon 类的构造函数
		/// </summary>
		public CustomRibbon()
		{
			_logger.LogInformation("构造 CustomRibbon");
			_tb101Press=false;
			// 注意：此时不初始化 _netApp，等待 SetApplication 调用
		}

		/// <summary>
		/// 在 ThisAddIn Startup 完成后设置 Application 对象
		/// </summary>
		/// <param name="application"> PowerPoint Application 实例 </param>
		public void SetApplication(NETOP.Application application)
		{
			if(application==null)
			{
				Logger.LogWarning("SetApplication 传入空 Application 对象");
				return;
			}

			_netApp=application;
			_appInitialized=true;

			// 从 DI 容器获取服务
			var serviceProvider = _applicationProvider?.ServiceProvider;
			if(serviceProvider!=null)
			{
				_logger=serviceProvider.GetService<ILogger>()??_logger;
			}

			Logger.LogInformation("Application 设置成功");
		}

		public void SetApplicationProvider(IApplicationProvider applicationProvider)
		{
			_applicationProvider=applicationProvider;
			_logger=applicationProvider?.ServiceProvider?.GetService<ILogger>()??_logger;

			// 初始化 Ribbon 相关服务
			var serviceProvider = applicationProvider?.ServiceProvider;
			if(serviceProvider!=null)
			{
				_ribbonXmlProvider=serviceProvider.GetService<IRibbonXmlProvider>();
				_ribbonIconProvider=serviceProvider.GetService<IRibbonIconProvider>();
				var shapeBatchHelper = serviceProvider.GetService<IShapeBatchHelper>();
				_selectionService=serviceProvider.GetService<ISelectionService>(); // 新增

				// 创建命令路由器（需要回调函数，所以在这里创建）
				_ribbonCommandRouter=new RibbonCommandRouter(
					serviceProvider,
					_logger,
					shapeBatchHelper,
					() => GetNetOfficeApplication(),
					_selectionService,
					() => _tb101Press,
					value => _tb101Press=value,
					controlId => _ribbonUI?.InvalidateControl(controlId),
					() => _ribbonUI?.Invalidate()
				);
			}
		}

		/// <summary>
		/// Ribbon UI 加载时调用的事件处理器
		/// </summary>
		/// <param name="ribbonUI"> 功能区UI接口 </param>
		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			try
			{
				_ribbonUI=ribbonUI;
				_ribbonIconProvider?.PreloadIcons();
				_tb101Press=false;

				_ribbonUI?.Invalidate();
				Logger.LogInformation("UI加载成功");
			} catch(Exception ex)
			{
				Logger.LogError($"UI加载错误: {ex.Message}",ex);
			}
		}

		/// <summary>
		/// IRibbonExtensibility 接口的实现，用于加载 Ribbon XML
		/// </summary>
		/// <param name="ribbonID"> 功能区标识符 </param>
		/// <returns> Ribbon的XML字符串 </returns>
		public string GetCustomUI(string ribbonID)
		{
			if(_ribbonXmlProvider!=null)
			{
				return _ribbonXmlProvider.GetRibbonXml(ribbonID);
			}

			// 后备方案：使用嵌入的资源字符串
			return Resources.RibbonXml;
		}

		#endregion Initialization & Setup

		#region State & Property Getters

		/// <summary>
		/// 获取 Ribbon 控件的图标
		/// </summary>
		public Bitmap GetIcon(Office.IRibbonControl control)
		{
			if(_ribbonIconProvider!=null)
			{
				bool? pressed = control.Id == "Tb101" ? _tb101Press : (bool?)null;
				return _ribbonIconProvider.GetIcon(control,pressed);
			}

			Logger.LogWarning("IRibbonIconProvider 未初始化");
			return null;
		}

		/// <summary>
		/// 获取切换按钮的标签
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		/// <returns> 切换按钮的显示文本 </returns>
		public string GetTbLabel(Office.IRibbonControl control)
		{
			// Profiler.LogMessage($"获取切换按钮标签 | {control.Id}");

			return control.Id switch
			{
				"Tb101" => _tb101Press
					? ResourceManager.GetString("Ribbon_Tb101_Slide","幻灯片")
					: ResourceManager.GetString("Ribbon_Tb101_Objects","所选对象"),
				_ => string.Empty,
			};
		}

		/// <summary>
		/// 获取 Ribbon 控件的标签文本（用于动态本地化）
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		/// <returns> 本地化的标签文本 </returns>
		public string GetLabel(Office.IRibbonControl control)
		{
			// 根据控件 ID 返回本地化字符串
			string resourceKey = $"Ribbon_{control.Id}";
			string defaultText = GetDefaultLabel(control.Id);
			return ResourceManager.GetString(resourceKey,defaultText);
		}

		/// <summary>
		/// 获取默认标签文本（当资源文件中找不到时使用）
		/// </summary>
		private string GetDefaultLabel(string controlId)
		{
			return controlId switch
			{
				"CustomTabXml" => "PPA菜单",
				"group1" => "对齐",
				"group11" => "吸附",
				"group2" => "大小",
				"group3" => "参考线",
				"group4" => "选择",
				"group5" => "格式",
				"group6" => "设置",
				"Bt101" => "左对齐",
				"Bt102" => "水平居中",
				"Bt103" => "右对齐",
				"Bt104" => "横向分布",
				"Bt111" => "顶对齐",
				"Bt112" => "垂直居中",
				"Bt113" => "底对齐",
				"Bt114" => "纵向分布",
				"Bt121" => "左吸附",
				"Bt122" => "右吸附",
				"Bt123" => "上吸附",
				"Bt124" => "下吸附",
				"Bt201" => "等宽度",
				"Bt202" => "等高度",
				"Bt203" => "等大小",
				"Bt204" => "互　换",
				"Bt211" => "左延伸",
				"Bt212" => "右延伸",
				"Bt213" => "上延伸",
				"Bt214" => "下延伸",
				"Bt301" => "左对齐",
				"Bt302" => "水平居中",
				"Bt303" => "右对齐",
				"Bt311" => "顶对齐",
				"Bt312" => "垂直居中",
				"Bt313" => "底对齐",
				"Bt321" => "宽扩展",
				"Bt322" => "高扩展",
				"Bt323" => "宽高扩展",
				"Bt401" => "隐显对象",
				"Bt402" => "裁剪出框",
				"Bt501" => "美化表格",
				"Bt502" => "美化文本",
				"Bt503" => "美化图表",
				"Bt601" => "插入形状",
				"MenuSettings" => "设置",
				"MenuLang_zhCN" => "中文 (简体)",
				"MenuLang_enUS" => "English (US)",
				"MenuSettings_Config" => "设置参数",
				"MenuSettings_About" => "关于",
				_ => string.Empty,
			};
		}

		/// <summary>
		/// 获取 NetOffice Application 对象（统一依赖 ApplicationHelper）
		/// </summary>
		/// <returns> NetOffice Application 对象，如果无法获取则返回 null </returns>
		private NETOP.Application GetNetOfficeApplication()
		{
			// 委托给 ApplicationHelper 处理获取和有效性验证（含自动重连逻辑）
			var validApp = ApplicationHelper.EnsureValidNetApplication(_netApp);

			if(validApp!=null)
			{
				_netApp=validApp;
				_appInitialized=true;
			} else
			{
				// 如果 EnsureValidNetApplication 返回 null，说明确实无法获取到有效的 App
				_appInitialized=false;
			}

			return _netApp;
		}

		/// <summary>
		/// 获取当前选中的形状数量
		/// </summary>
		/// <returns> 选中的形状数量，如果无法获取则返回 0 </returns>
		private int GetSelectedShapeCount()
		{
			return _selectionService?.GetSelectedShapeCount()??0;
		}

		/// <summary>
		/// 获取切换按钮的按下状态
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		/// <returns> 切换按钮的当前状态 </returns>
		public bool GetTbState(Office.IRibbonControl control)
		{
			if(control.Id=="Tb101")
			{
				int currentShapeCount = GetSelectedShapeCount();
				if(currentShapeCount<=1)
				{
					// 单个形状或未选：强制对齐到幻灯片（ObjectsAlignRelativeToContainerSmart）
					_tb101Press=true;
				}
				// 检测选中数量是否变化，如果变化则刷新 UI
				if(currentShapeCount!=_lastShapeCount)
				{
					int previousShapeCount = _lastShapeCount;
					_lastShapeCount=currentShapeCount;
					// 异步刷新 Ribbon UI，避免在回调中直接刷新导致的问题
					ThreadPool.QueueUserWorkItem(_ =>
					{
						try
						{
							Thread.Sleep(50); // 短暂延迟，确保状态已更新
							_ribbonUI?.InvalidateControl("Tb101");
						} catch(Exception ex)
						{
							Logger.LogError($"刷新 Tb101 按钮状态时出错: {ex.Message}",ex);
						}
					});

					Logger.LogDebug($"获取切换按钮状态: {control.Id}, 选中形状数: {currentShapeCount}, 状态: {(_tb101Press ? "幻灯片" : "所选对象")}");
				}
			}

			return control.Id switch
			{
				"Tb101" => _tb101Press,
				_ => false,
			};
		}

		#endregion State & Property Getters

		#region Event Handlers

		/// <summary>
		/// 处理普通按钮的点击事件
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		public void OnAction(Office.IRibbonControl control)
		{
			Logger.LogInformation($"按钮点击事件: {control.Id}");

			if(!_appInitialized||_netApp==null)
			{
				Logger.LogWarning($"Application 未初始化，跳过操作: {control.Id}");
				return;
			}

			if(_ribbonCommandRouter!=null)
			{
				_ribbonCommandRouter.ExecuteButtonCommand(control.Id);
			} else
			{
				Logger.LogWarning("IRibbonCommandRouter 未初始化，无法执行命令");
			}
		}

		/// <summary>
		/// 处理切换按钮的点击事件
		/// </summary>
		public void TbOnAction(Office.IRibbonControl control,bool pressed)
		{
			if(_ribbonCommandRouter!=null)
			{
				_ribbonCommandRouter.HandleToggleButton(control,pressed);
			} else
			{
				Logger.LogWarning("IRibbonCommandRouter 未初始化，无法处理切换按钮");
			}
		}

		/// <summary>
		/// 处理菜单项的点击事件
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		public void OnMenuAction(Office.IRibbonControl control)
		{
			Logger.LogInformation($"菜单项点击事件: {control.Id}");

			if(_ribbonCommandRouter!=null)
			{
				_ribbonCommandRouter.HandleMenuAction(control);
			} else
			{
				Logger.LogWarning("IRibbonCommandRouter 未初始化，无法处理菜单项");
			}
		}

		/// <summary>
		/// 获取菜单项图标（用于语言选择标记）
		/// </summary>
		public Bitmap GetMenuIcon(Office.IRibbonControl control)
		{
			// 为当前选中的语言显示标记
			if(control.Id=="MenuLang_zhCN"&&ResourceManager.CurrentCulture.Name=="zh-CN")
			{
				return CreateCheckIcon();
			}
			if(control.Id=="MenuLang_enUS"&&ResourceManager.CurrentCulture.Name=="en-US")
			{
				return CreateCheckIcon();
			}
			return null;
		}

		/// <summary>
		/// 创建选中标记图标
		/// </summary>
		private Bitmap CreateCheckIcon()
		{
			var bmp = new Bitmap(16, 16);
			using(var g = Graphics.FromImage(bmp))
			{
				g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
				using var pen = new Pen(Color.Green,2);
				// 绘制对勾
				g.DrawLine(pen,3,8,7,12);
				g.DrawLine(pen,7,12,13,4);
			}
			return bmp;
		}

		/// <summary>
		/// 显示设置对话框
		/// </summary>
		private void ShowSettingsDialog()
		{
			try
			{
				using var settingsForm = new SettingsForm();
				settingsForm.ShowDialog();
			} catch(Exception ex)
			{
				Logger.LogError($"显示设置对话框失败: {ex.Message}",ex);
				Toast.Show($"打开设置窗口失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		/// <summary>
		/// 显示关于对话框
		/// </summary>
		private void ShowAboutDialog()
		{
			try
			{
				using var aboutForm = new AboutForm();
				aboutForm.ShowDialog();
			} catch(Exception ex)
			{
				Logger.LogError($"显示关于对话框失败: {ex.Message}",ex);
				Toast.Show($"打开关于窗口失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		#endregion Event Handlers

		#region Lifecycle Management (IDisposable)

		/// <summary>
		/// 公共的 Dispose 方法
		/// </summary>
		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// 受保护的 Dispose 方法，用于释放资源
		/// </summary>
		protected virtual void Dispose(bool disposing)
		{
			if(_disposed) return;

			if(disposing)
			{
				Logger.LogDebug("CustomRibbon 释放资源");

				// 释放图标资源
				_ribbonIconProvider?.DisposeIcons();

				try
				{
					if(_ribbonUI!=null)
					{
						Marshal.ReleaseComObject(_ribbonUI);
						_ribbonUI=null;
					}
				} catch(Exception ex)
				{
					Logger.LogWarning($"释放UI时出错: {ex.Message}");
				}

				// 注意：不释放 _netApp，因为它由 ThisAddIn 管理
			}

			_disposed=true;
		}

		#endregion Lifecycle Management (IDisposable)
	}
}

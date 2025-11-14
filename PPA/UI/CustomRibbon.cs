using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;
using PPA.Formatting;
using PPA.Properties;
using PPA.Shape;
using PPA.UI.Forms;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using ALT = PPA.Core.Abstraction.Business.AlignmentType;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA
{
	[ComVisible(true)]
	public class CustomRibbon:Office.IRibbonExtensibility, IDisposable
	{
		#region Private Fields

		private Office.IRibbonUI _ribbonUI;
		private NETOP.Application _app;
		private readonly Dictionary<string,Bitmap> _iconCache;
		private bool _tb101Press;
		private bool _disposed = false;
		private bool _appInitialized = false;
		private CancellationTokenSource _bt501Cancellation;
		private AlignHelper _alignHelper; // 对齐工具服务（使用具体类型以访问所有方法）
		private IApplication _abstractApp;

		#endregion Private Fields

		#region Initialization & Setup

		/// <summary>
		/// CustomRibbon 类的构造函数
		/// </summary>
		public CustomRibbon()
		{
			Profiler.LogMessage("构造...");
			_iconCache= [];
			_tb101Press=false;
			// 注意：此时不初始化 _app，等待 SetApplication 调用
		}

		/// <summary>
		/// 在 ThisAddIn Startup 完成后设置 Application 对象
		/// </summary>
		/// <param name="application"> PowerPoint Application 实例 </param>
		public void SetApplication(NETOP.Application application)
		{
			if(application==null)
			{
				Profiler.LogMessage("SetApplication 传入空 Application 对象");
				return;
			}

			_app=application;
			_appInitialized=true;

			// 从 DI 容器获取服务
			var addIn = Globals.ThisAddIn;
			if (addIn != null && addIn.ServiceProvider != null)
			{
				var service = addIn.ServiceProvider.GetService<IAlignHelper>();
				_alignHelper = service as AlignHelper;
			}

			_abstractApp = ResolveAbstractApplication();

			Profiler.LogMessage("Application 设置成功");
		}

		/// <summary>
		/// Ribbon UI 加载时调用的事件处理器
		/// </summary>
		/// <param name="ribbonUI"> 功能区UI接口 </param>
		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			try
			{
				// Profiler.LogMessage(UI初始化...");
				_ribbonUI=ribbonUI;
				PreloadIcons();
				_tb101Press=false;

				_ribbonUI?.Invalidate();
				Profiler.LogMessage("UI加载成功");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"UI加载错误: {ex.Message}");
			}
		}

		/// <summary>
		/// IRibbonExtensibility 接口的实现，用于加载 Ribbon XML
		/// </summary>
		/// <param name="ribbonID"> 功能区标识符 </param>
		/// <returns> Ribbon的XML字符串 </returns>
		public string GetCustomUI(string ribbonID)
		{
			// Profiler.LogMessage($"ribbonID: {ribbonID}");

			try
			{
				string ribbonXml = LoadRibbonXmlFromFile();
				if(!string.IsNullOrEmpty(ribbonXml))
				{
					return ribbonXml;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"从文件加载XML失败: {ex.Message}");
			}

			return Resources.RibbonXml;
		}

		/// <summary>
		/// 预加载所有 Ribbon 图标到缓存中，提高UI响应速度
		/// </summary>
		private void PreloadIcons()
		{
			if(_iconCache.Count>0) return;

			// Profiler.LogMessage("预加载图标...");

			try
			{
				Dictionary<string, Bitmap> icons = new()
				{
					["Tb101_1"] = Properties.Resources.slide,
					["Tb101_0"] = Properties.Resources.shap,
					["Bt121"] = Properties.Resources.Bt121,
					["Bt122"] = Properties.Resources.Bt122,
					["Bt123"] = Properties.Resources.Bt123,
					["Bt124"] = Properties.Resources.Bt124,
					["Bt204"] = Properties.Resources.Bt204,
					["Bt211"] = Properties.Resources.Bt211,
					["Bt212"] = Properties.Resources.Bt212,
					["Bt213"] = Properties.Resources.Bt213,
					["Bt214"] = Properties.Resources.Bt214,
					["Bt301"] = Properties.Resources.Bt301,
					["Bt302"] = Properties.Resources.Bt302,
					["Bt303"] = Properties.Resources.Bt303,
					["Bt311"] = Properties.Resources.Bt311,
					["Bt312"] = Properties.Resources.Bt312,
					["Bt313"] = Properties.Resources.Bt313,
					["Bt321"] = Properties.Resources.Bt321,
					["Bt322"] = Properties.Resources.Bt323,
					["Bt323"] = Properties.Resources.Bt322,
					["Bt401"] = Properties.Resources.Bt401,
					["Bt402"] = Properties.Resources.Bt402,
					["Bt601"] = Properties.Resources.Bt601
				};

				foreach(var icon in icons)
				{
					_iconCache[icon.Key]=icon.Value;
				}

				// Profiler.LogMessage($"已预加载 {_iconCache.Count} 个图标");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"预加载图标错误: {ex.Message}");
			}
		}

		#endregion Initialization & Setup

		#region State & Property Getters

		/// <summary>
		/// 获取 Ribbon 控件的图标
		/// </summary>
		public Bitmap GetIcon(Office.IRibbonControl control)
		{
			try
			{
				string itemId = control.Id;
				if(control.Id=="Tb101")
				{
					itemId=_tb101Press ? "Tb101_1" : "Tb101_0";
				}

				if(_iconCache.TryGetValue(itemId,out Bitmap bmp))
				{
					return bmp;
				}

				Profiler.LogMessage($"未找到图标: {itemId}");
				return null;
			} catch(Exception ex)
			{
				Profiler.LogMessage($"获取图标错误 | {control.Id}: {ex.Message}");
				return null;
			}
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
					? ResourceManager.GetString("Ribbon_Tb101_Page","页面")
					: ResourceManager.GetString("Ribbon_Tb101_Shape","形状"),
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
		/// 获取切换按钮的按下状态
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		/// <returns> 切换按钮的当前状态 </returns>
		public bool GetTbState(Office.IRibbonControl control)
		{
			Profiler.LogMessage($"获取切换按钮状态: {control.Id}");

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
			Profiler.LogMessage($"按钮点击事件: {control.Id}");

			if(!_appInitialized||_app==null)
			{
				Profiler.LogMessage($"Application 未初始化，跳过操作: {control.Id}");
				return;
			}

			try
			{
				ExecuteButtonAction(control.Id);
			} catch(Exception ex)
			{
				Profiler.LogMessage($"按钮操作错误 {control.Id}: {ex.Message}");
			}
		}

		/// <summary>
		/// 处理切换按钮的点击事件
		/// </summary>
		public void TbOnAction(Office.IRibbonControl control,bool pressed)
		{
			// Profiler.LogMessage($"切换按钮点击事件 | {control.Id}, pressed: {pressed}");

			try
			{
				if(control.Id=="Tb101")
				{
					_tb101Press=pressed;
					_ribbonUI?.InvalidateControl("Tb101");
					Profiler.LogMessage($"切换状态 | {control.Id}: {(_tb101Press ? "页面" : "形状")}");
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"切换按钮点击事件错误 | {control.Id}: {ex.Message}");
			}
		}

		/// <summary>
		/// 处理菜单项的点击事件
		/// </summary>
		/// <param name="control"> 功能区控件对象 </param>
		public void OnMenuAction(Office.IRibbonControl control)
		{
			Profiler.LogMessage($"菜单项点击事件: {control.Id}");

			try
			{
				switch(control.Id)
				{
					case "MenuLang_zhCN":
					{
						bool ok = ResourceManager.SetLanguage("zh-CN");
						if (ok)
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChanged","语言已切换为中文"),Toast.ToastType.Success);
							// 刷新整个 Ribbon 以更新所有文本
							_ribbonUI?.Invalidate();
						}
						else
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChangeFailed","切换语言失败"), Toast.ToastType.Error);
						}
						break;
					}

					case "MenuLang_enUS":
					{
						bool ok = ResourceManager.SetLanguage("en-US");
						if (ok)
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChanged","Language switched to English"),Toast.ToastType.Success);
							// 刷新整个 Ribbon 以更新所有文本
							_ribbonUI?.Invalidate();
						}
						else
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChangeFailed","Language change failed"), Toast.ToastType.Error);
						}
						break;
					}

					case "MenuSettings_Config": ShowSettingsDialog(); break;
					case "MenuSettings_About": ShowAboutDialog(); break;
					default: Profiler.LogMessage($"未知菜单项ID: {control.Id}"); break;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"菜单项操作错误 {control.Id}: {ex.Message}");
				Toast.Show($"操作失败: {ex.Message}",Toast.ToastType.Error);
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
				Profiler.LogMessage($"显示设置对话框失败: {ex.Message}");
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
				Profiler.LogMessage($"显示关于对话框失败: {ex.Message}");
				Toast.Show($"打开关于窗口失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		#endregion Event Handlers

		#region Core Business Logic

		/// <summary>
		/// 根据按钮 ID 执行相应的业务逻辑
		/// </summary>
		/// <param name="buttonId"> 按钮标识符 </param>
		private void ExecuteButtonAction(string buttonId)
		{
			if(!_appInitialized||_app==null)
			{
				Profiler.LogMessage($"Application 不可用，无法执行操作");
				return;
			}

			// 如果未获取到服务，尝试从 DI 容器获取（向后兼容）
			if (_alignHelper == null)
			{
				var addIn = Globals.ThisAddIn;
				if (addIn != null && addIn.ServiceProvider != null)
				{
					var service = addIn.ServiceProvider.GetService<IAlignHelper>();
					_alignHelper = service as AlignHelper;
				}
			}

			if (_alignHelper == null)
			{
				Profiler.LogMessage("警告：无法获取 IAlignHelper 服务，创建新实例");
				// 向后兼容：如果无法获取服务，创建临时实例
				_alignHelper = new AlignHelper();
			}

			switch(buttonId)
			{
				case "Bt101":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Left, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Left, _tb101Press));
					break;
				case "Bt102":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Centers, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Centers, _tb101Press));
					break;
				case "Bt103":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Right, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Right, _tb101Press));
					break;
				case "Bt104":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Horizontally, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Horizontally, _tb101Press));
					break;
				case "Bt111":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Top, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Top, _tb101Press));
					break;
				case "Bt112":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Middles, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Middles, _tb101Press));
					break;
				case "Bt113":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Bottom, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Bottom, _tb101Press));
					break;
				case "Bt114":
					PerformAlignment(
						(helper, app) => helper.ExecuteAlignment(app, ALT.Vertically, _tb101Press),
						(helper, app) => helper.ExecuteAlignment(app, ALT.Vertically, _tb101Press));
					break;
				case "Bt121":
					PerformAlignment(
						(helper, app) => helper.AttachLeft(app),
						(helper, app) => helper.AttachLeft(app));
					break;
				case "Bt122":
					PerformAlignment(
						(helper, app) => helper.AttachRight(app),
						(helper, app) => helper.AttachRight(app));
					break;
				case "Bt123":
					PerformAlignment(
						(helper, app) => helper.AttachTop(app),
						(helper, app) => helper.AttachTop(app));
					break;
				case "Bt124":
					PerformAlignment(
						(helper, app) => helper.AttachBottom(app),
						(helper, app) => helper.AttachBottom(app));
					break;
				case "Bt201":
					PerformAlignment(
						(helper, app) => helper.SetEqualWidth(app),
						(helper, app) => helper.SetEqualWidth(app));
					break;
				case "Bt202":
					PerformAlignment(
						(helper, app) => helper.SetEqualHeight(app),
						(helper, app) => helper.SetEqualHeight(app));
					break;
				case "Bt203":
					PerformAlignment(
						(helper, app) => helper.SetEqualSize(app),
						(helper, app) => helper.SetEqualSize(app));
					break;
				case "Bt204":
					PerformAlignment(
						(helper, app) => helper.SwapSize(app),
						(helper, app) => helper.SwapSize(app));
					break;
				case "Bt211":
					PerformAlignment(
						(helper, app) => helper.StretchLeft(app),
						(helper, app) => helper.StretchLeft(app));
					break;
				case "Bt212":
					PerformAlignment(
						(helper, app) => helper.StretchRight(app),
						(helper, app) => helper.StretchRight(app));
					break;
				case "Bt213":
					PerformAlignment(
						(helper, app) => helper.StretchTop(app),
						(helper, app) => helper.StretchTop(app));
					break;
				case "Bt214":
					PerformAlignment(
						(helper, app) => helper.StretchBottom(app),
						(helper, app) => helper.StretchBottom(app));
					break;
				case "Bt301":
					PerformAlignment(
						(helper, app) => helper.GuideAlignLeft(app),
						(helper, app) => helper.GuideAlignLeft(app));
					break;
				case "Bt302":
					PerformAlignment(
						(helper, app) => helper.GuideAlignHCenter(app),
						(helper, app) => helper.GuideAlignHCenter(app));
					break;
				case "Bt303":
					PerformAlignment(
						(helper, app) => helper.GuideAlignRight(app),
						(helper, app) => helper.GuideAlignRight(app));
					break;
				case "Bt311":
					PerformAlignment(
						(helper, app) => helper.GuideAlignTop(app),
						(helper, app) => helper.GuideAlignTop(app));
					break;
				case "Bt312":
					PerformAlignment(
						(helper, app) => helper.GuideAlignVCenter(app),
						(helper, app) => helper.GuideAlignVCenter(app));
					break;
				case "Bt313":
					PerformAlignment(
						(helper, app) => helper.GuideAlignBottom(app),
						(helper, app) => helper.GuideAlignBottom(app));
					break;
				case "Bt321":
					PerformAlignment(
						(helper, app) => helper.GuidesStretchWidth(app),
						(helper, app) => helper.GuidesStretchWidth(app));
					break;
				case "Bt322":
					PerformAlignment(
						(helper, app) => helper.GuidesStretchHeight(app),
						(helper, app) => helper.GuidesStretchHeight(app));
					break;
				case "Bt323":
					PerformAlignment(
						(helper, app) => helper.GuidesStretchSize(app),
						(helper, app) => helper.GuidesStretchSize(app));
					break;
				case "Bt401": ShapeBatchHelper.ToggleShapeVisibility(_app); break;
				case "Bt402": MSOICrop.CropShapesToSlide(_app); break;

				case "Bt501":
					// 取消之前的操作（如果存在）
					_bt501Cancellation?.Cancel();
					_bt501Cancellation?.Dispose();
					_bt501Cancellation=new CancellationTokenSource();

					// 异步执行美化表格（fire-and-forget，异常已在 ExecuteAsyncOperation 内部处理）
					AsyncOperationHelper.ExecuteAsyncOperation(async () =>
					{
						var progress = new ProgressIndicator(ResourceManager.GetString("Ribbon_Bt501", "美化表格"));
						var helper = ResolveTableBatchHelper();
						if(helper!=null)
						{
							await helper.FormatTablesAsync(
								_app,
								progress,
								_bt501Cancellation.Token);
						}
						else
						{
							await TableBatchHelper.Bt501_ClickAsync(
								_app,
								progress,
								_bt501Cancellation.Token);
						}
					},ResourceManager.GetString("Ribbon_Bt501","美化表格"));
					// 注意：ExecuteAsyncOperation 是 async void，设计为 fire-and-forget，不需要 await
					break;

				case "Bt502":
				{
					var helper = ResolveTextBatchHelper();
					if(helper != null)
					{
						helper.FormatText(_app);
					}
					else
					{
						TextBatchHelper.Bt502_Click(_app);
					}
					break;
				}
				case "Bt503":
				{
					var helper = ResolveChartBatchHelper();
					if(helper!=null)
					{
						helper.FormatCharts(_app);
					}
					else
					{
						ChartBatchHelper.Bt503_Click(_app);
					}
					break;
				}
				case "Bt601": ShapeBatchHelper.Bt601_Click(_app); break;
				default: Profiler.LogMessage($"未知按钮ID: {buttonId}"); break;
			}
		}

		#endregion Core Business Logic

		#region Async Operation Helpers

		// ExecuteAsyncOperation 已移动到 AsyncOperationHelper 类中

		#endregion Async Operation Helpers

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
				Profiler.LogMessage($"释放资源...");

				// 取消并清理异步操作
				_bt501Cancellation?.Cancel();
				_bt501Cancellation?.Dispose();
				_bt501Cancellation=null;

				foreach(var kvp in _iconCache)
				{
					try
					{
						kvp.Value?.Dispose();
					} catch(Exception ex)
					{
						Profiler.LogMessage($"释放图标资源时出错 | {kvp.Key}: {ex.Message}");
					}
				}
				_iconCache.Clear();

				try
				{
					if(_ribbonUI!=null)
					{
						Marshal.ReleaseComObject(_ribbonUI);
						_ribbonUI=null;
					}
				} catch(Exception ex)
				{
					Profiler.LogMessage($"释放UI时出错: {ex.Message}");
				}

				// 注意：不释放 _app，因为它由 ThisAddIn 管理
			}

			_disposed=true;
		}

		#endregion Lifecycle Management (IDisposable)

		#region Private Helper Methods

		/// <summary>
		/// 从嵌入式资源中加载 Ribbon XML，如果找不到则返回 null
		/// 使用 .NET Framework 的资源加载机制，自动处理 ClickOnce 部署
		/// </summary>
		/// <returns> 加载的XML字符串，如未找到则返回null </returns>
		private string LoadRibbonXmlFromFile()
		{
			try
			{
				// 从嵌入式资源加载 Ribbon.xml
				// 资源名称格式：命名空间.文件夹.文件名
				string resourceName = "PPA.UI.Ribbon.xml";
				var assembly = System.Reflection.Assembly.GetExecutingAssembly();
				
				using(var stream = assembly.GetManifestResourceStream(resourceName))
				{
					if(stream != null)
					{
						using var reader = new StreamReader(stream);
						string xmlContent = reader.ReadToEnd();
						Profiler.LogMessage($"成功从嵌入式资源加载 Ribbon.xml");
						return xmlContent;
					}
				}
				
				Profiler.LogMessage($"未找到嵌入式资源: {resourceName}，使用后备资源");
			}
			catch(Exception ex)
			{
				Profiler.LogMessage($"从嵌入式资源加载 Ribbon.xml 失败: {ex.Message}");
			}
			
			return null;
		}

		#endregion Private Helper Methods

		private ITextBatchHelper ResolveTextBatchHelper()
		{
			var addIn = Globals.ThisAddIn;
			return addIn?.ServiceProvider?.GetService<ITextBatchHelper>();
		}

		private IChartBatchHelper ResolveChartBatchHelper()
		{
			var addIn = Globals.ThisAddIn;
			return addIn?.ServiceProvider?.GetService<IChartBatchHelper>();
		}

		private ITableBatchHelper ResolveTableBatchHelper()
		{
			var addIn = Globals.ThisAddIn;
			return addIn?.ServiceProvider?.GetService<ITableBatchHelper>();
		}

	private IApplication GetAbstractApplication()
	{
		if(_abstractApp == null)
		{
			_abstractApp = ResolveAbstractApplication();
		}
		return _abstractApp;
	}

	private IApplication ResolveAbstractApplication()
	{
		var addIn = Globals.ThisAddIn;
		var serviceProvider = addIn?.ServiceProvider;
		if(serviceProvider != null)
		{
			var factory = serviceProvider.GetService<IApplicationFactory>();
			var app = factory?.GetCurrent();
			if(app != null)
			{
				return app;
			}
		}

		return _app != null ? new PowerPointApplication(_app) : null;
	}

	private void PerformAlignment(Action<AlignHelper, IApplication> abstractAction, Action<AlignHelper, NETOP.Application> nativeAction)
	{
		var abstractApp = GetAbstractApplication();
		if(abstractApp != null)
		{
			abstractAction?.Invoke(_alignHelper, abstractApp);
		}
		else
		{
			nativeAction?.Invoke(_alignHelper, _app);
		}
	}
	}
}

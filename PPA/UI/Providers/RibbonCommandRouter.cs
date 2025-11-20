using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;
using PPA.Core.Logging;
using PPA.Formatting;
using PPA.Shape;
using PPA.UI.Forms;
using PPA.Utilities;
using System;
using ALT = PPA.Core.Abstraction.Business.AlignmentType;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA.UI.Providers
{
	/// <summary>
	/// Ribbon 命令路由实现
	/// </summary>
	internal sealed class RibbonCommandRouter:IRibbonCommandRouter
	{
		private readonly IServiceProvider _serviceProvider;
		private readonly IApplicationProvider _applicationProvider;
		private readonly ILogger _logger;
		private readonly IShapeBatchHelper _shapeBatchHelper;
		private readonly Func<NETOP.Application> _getNetApp;
		private readonly Func<IApplication> _getAbstractApp;
		private readonly Func<bool> _getTb101Press;
		private readonly Action<bool> _setTb101Press;
		private readonly Func<int> _getSelectedShapeCount;
		private readonly Action<string> _invalidateControl;
		private readonly Action _invalidateRibbon;

		public RibbonCommandRouter(
			IServiceProvider serviceProvider,
			IApplicationProvider applicationProvider,
			ILogger logger,
			IShapeBatchHelper shapeBatchHelper,
			Func<NETOP.Application> getNetApp,
			Func<IApplication> getAbstractApp,
			Func<bool> getTb101Press,
			Action<bool> setTb101Press,
			Func<int> getSelectedShapeCount,
			Action<string> invalidateControl,
			Action invalidateRibbon)
		{
			_serviceProvider=serviceProvider??throw new ArgumentNullException(nameof(serviceProvider));
			_applicationProvider=applicationProvider??throw new ArgumentNullException(nameof(applicationProvider));
			_logger=logger??LoggerProvider.GetLogger();
			_shapeBatchHelper=shapeBatchHelper??throw new ArgumentNullException(nameof(shapeBatchHelper));
			_getNetApp=getNetApp??throw new ArgumentNullException(nameof(getNetApp));
			_getAbstractApp=getAbstractApp??throw new ArgumentNullException(nameof(getAbstractApp));
			_getTb101Press=getTb101Press??throw new ArgumentNullException(nameof(getTb101Press));
			_setTb101Press=setTb101Press??throw new ArgumentNullException(nameof(setTb101Press));
			_getSelectedShapeCount=getSelectedShapeCount??throw new ArgumentNullException(nameof(getSelectedShapeCount));
			_invalidateControl=invalidateControl??throw new ArgumentNullException(nameof(invalidateControl));
			_invalidateRibbon=invalidateRibbon??throw new ArgumentNullException(nameof(invalidateRibbon));
		}

		/// <summary>
		/// 执行按钮命令
		/// </summary>
		public bool ExecuteButtonCommand(string buttonId)
		{
			var netApp = _getNetApp();
			if(netApp==null)
			{
				_logger.LogWarning("Application 不可用，无法执行操作");
				return false;
			}

			// Ribbon 层不再检查 ActiveWindow，交由业务逻辑决定是否可执行

			// 获取对齐助手
			var alignHelper = ResolveAlignHelper();
			if(alignHelper==null)
			{
				_logger.LogWarning("无法获取 IAlignHelper 服务");
				return false;
			}

			// 在执行对齐操作前刷新切换按钮 UI
			if(buttonId.StartsWith("Bt10")||buttonId.StartsWith("Bt11"))
			{
				_invalidateControl("Tb101");
			}

			bool tb101Press = _getTb101Press();
			var abstractApp = _getAbstractApp();

			try
			{
				switch(buttonId)
				{
					case "Bt101":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Left,tb101Press));
						return true;

					case "Bt102":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Centers,tb101Press));
						return true;

					case "Bt103":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Right,tb101Press));
						return true;

					case "Bt104":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Horizontally,tb101Press));
						return true;

					case "Bt111":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Top,tb101Press));
						return true;

					case "Bt112":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Middles,tb101Press));
						return true;

					case "Bt113":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Bottom,tb101Press));
						return true;

					case "Bt114":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.ExecuteAlignment(a,ALT.Vertically,tb101Press));
						return true;

					case "Bt121":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.AttachLeft(a));
						return true;

					case "Bt122":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.AttachRight(a));
						return true;

					case "Bt123":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.AttachTop(a));
						return true;

					case "Bt124":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.AttachBottom(a));
						return true;

					case "Bt201":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.SetEqualWidth(a));
						return true;

					case "Bt202":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.SetEqualHeight(a));
						return true;

					case "Bt203":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.SetEqualSize(a));
						return true;

					case "Bt204":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.SwapSize(a));
						return true;

					case "Bt211":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.StretchLeft(a));
						return true;

					case "Bt212":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.StretchRight(a));
						return true;

					case "Bt213":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.StretchTop(a));
						return true;

					case "Bt214":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.StretchBottom(a));
						return true;

					case "Bt301":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignLeft(a));
						return true;

					case "Bt302":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignHCenter(a));
						return true;

					case "Bt303":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignRight(a));
						return true;

					case "Bt311":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignTop(a));
						return true;

					case "Bt312":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignVCenter(a));
						return true;

					case "Bt313":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuideAlignBottom(a));
						return true;

					case "Bt321":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuidesStretchWidth(a));
						return true;

					case "Bt322":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuidesStretchHeight(a));
						return true;

					case "Bt323":
						PerformAlignment(alignHelper,abstractApp,(h,a) => h.GuidesStretchSize(a));
						return true;

					case "Bt401":
						_shapeBatchHelper.ToggleShapeVisibility(netApp);
						return true;

					case "Bt402":
					{
						var validApp = ApplicationHelper.EnsureValidNetApplication(netApp);
						if(validApp!=null)
						{
							MSOICrop.CropShapesToSlide(validApp);
						} else
						{
							_logger.LogWarning("Bt402: 无法获取有效的 Application");
						}
						return true;
					}

					case "Bt501":
					{
						var helper = ResolveTableBatchHelper();
						if(helper==null)
						{
							_logger.LogWarning("无法获取 ITableBatchHelper 服务");
							return true;
						}
						helper.FormatTables(netApp);
						return true;
					}

					case "Bt502":
					{
						var helper = ResolveTextBatchHelper();
						if(helper==null)
						{
							_logger.LogWarning("无法获取 ITextBatchHelper 服务");
							return false;
						}
						helper.FormatText(netApp);
						return true;
					}

					case "Bt503":
					{
						var helper = ResolveChartBatchHelper();
						if(helper==null)
						{
							_logger.LogWarning("无法获取 IChartBatchHelper 服务");
							return false;
						}
						helper.FormatCharts(netApp);
						return true;
					}

					case "Bt601":
						_shapeBatchHelper.CreateBoundingBox(netApp);
						return true;

					default:
						_logger.LogWarning($"未知按钮ID: {buttonId}");
						return false;
				}
			} catch(Exception ex)
			{
				_logger.LogError($"执行按钮命令失败 {buttonId}: {ex.Message}",ex);
				return false;
			}
		}

		/// <summary>
		/// 处理切换按钮的点击事件
		/// </summary>
		public bool HandleToggleButton(Office.IRibbonControl control,bool pressed)
		{
			if(control.Id!="Tb101")
			{
				return false;
			}

			try
			{
				int shapeCount = _getSelectedShapeCount();
				var commandExecutor = _serviceProvider.GetService<ICommandExecutor>();

				if(shapeCount>=2)
				{
					// 大于等于2个对象：切换状态
					bool previousState = _getTb101Press();
					_setTb101Press(pressed);

					if(previousState!=pressed&&commandExecutor!=null)
					{
						string msoCommand = pressed
							? OfficeCommands.ObjectsAlignRelativeToContainerSmart
							: OfficeCommands.ObjectsAlignSelectedSmart;

						bool success = commandExecutor.ExecuteMso(msoCommand);
						if(success)
						{
							_logger.LogInformation($"切换状态并执行 MSO 命令 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 命令: {msoCommand}, 选中形状数: {shapeCount}");
						} else
						{
							_logger.LogWarning($"切换状态但 MSO 命令执行失败 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 命令: {msoCommand}, 选中形状数: {shapeCount}");
						}
					} else if(previousState==pressed)
					{
						_logger.LogDebug($"状态未变化，跳过执行 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 选中形状数: {shapeCount}");
					} else
					{
						_logger.LogWarning($"切换状态 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 选中形状数: {shapeCount} (命令执行器不可用)");
					}

					_invalidateControl("Tb101");
				} else
				{
					// 小于2个对象：设置为对齐幻灯片
					bool previousState = _getTb101Press();
					_setTb101Press(true);

					if(!previousState&&commandExecutor!=null)
					{
						bool success = commandExecutor.ExecuteMso(OfficeCommands.ObjectsAlignRelativeToContainerSmart);
						if(!success)
						{
							_logger.LogWarning($"设置为对齐幻灯片但 MSO 命令执行失败 | {control.Id}: 命令: {OfficeCommands.ObjectsAlignRelativeToContainerSmart}, 选中形状数: {shapeCount}");
						}
					}
				}

				_invalidateControl("Tb101");
				return true;
			} catch(Exception ex)
			{
				_logger.LogError($"切换按钮点击事件错误 | {control.Id}: {ex.Message}",ex);
				return false;
			}
		}

		/// <summary>
		/// 处理菜单项的点击事件
		/// </summary>
		public bool HandleMenuAction(Office.IRibbonControl control)
		{
			try
			{
				switch(control.Id)
				{
					case "MenuLang_zhCN":
					{
						bool ok = ResourceManager.SetLanguage("zh-CN");
						if(ok)
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChanged","语言已切换为中文"),Toast.ToastType.Success);
							_invalidateRibbon();
						} else
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChangeFailed","切换语言失败"),Toast.ToastType.Error);
						}
						return true;
					}

					case "MenuLang_enUS":
					{
						bool ok = ResourceManager.SetLanguage("en-US");
						if(ok)
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChanged","Language switched to English"),Toast.ToastType.Success);
							_invalidateRibbon();
						} else
						{
							Toast.Show(ResourceManager.GetString("Settings_LanguageChangeFailed","Language change failed"),Toast.ToastType.Error);
						}
						return true;
					}

					case "MenuSettings_Config":
						ShowSettingsDialog();
						return true;

					case "MenuSettings_About":
						ShowAboutDialog();
						return true;

					default:
						_logger.LogWarning($"未知菜单项ID: {control.Id}");
						return false;
				}
			} catch(Exception ex)
			{
				_logger.LogError($"菜单项操作错误 {control.Id}: {ex.Message}",ex);
				Toast.Show($"操作失败: {ex.Message}",Toast.ToastType.Error);
				return false;
			}
		}

		#region Private Helper Methods

		private AlignHelper ResolveAlignHelper()
		{
			var service = _serviceProvider.GetService<IAlignHelper>();
			if(service is AlignHelper alignHelper)
			{
				return alignHelper;
			}
			return new AlignHelper();
		}

		private ITextBatchHelper ResolveTextBatchHelper()
		{
			return _serviceProvider.GetService<ITextBatchHelper>();
		}

		private IChartBatchHelper ResolveChartBatchHelper()
		{
			return _serviceProvider.GetService<IChartBatchHelper>();
		}

		private ITableBatchHelper ResolveTableBatchHelper()
		{
			return _serviceProvider.GetService<ITableBatchHelper>();
		}

		private void PerformAlignment(AlignHelper helper,IApplication abstractApp,Action<AlignHelper,IApplication> action)
		{
			if(abstractApp!=null)
			{
				action?.Invoke(helper,abstractApp);
			} else
			{
				var netApp = _getNetApp();
				if(netApp!=null)
				{
					var wrappedApp = new PowerPointApplication(netApp);
					action?.Invoke(helper,wrappedApp);
				}
			}
		}

		private void ShowSettingsDialog()
		{
			try
			{
				using var settingsForm = new SettingsForm();
				settingsForm.ShowDialog();
			} catch(Exception ex)
			{
				_logger.LogError($"显示设置对话框失败: {ex.Message}",ex);
				Toast.Show($"打开设置窗口失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		private void ShowAboutDialog()
		{
			try
			{
				using var aboutForm = new AboutForm();
				aboutForm.ShowDialog();
			} catch(Exception ex)
			{
				_logger.LogError($"显示关于对话框失败: {ex.Message}",ex);
				Toast.Show($"打开关于窗口失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		#endregion Private Helper Methods
	}
}

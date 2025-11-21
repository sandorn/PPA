using Microsoft.Extensions.DependencyInjection;
using NetOffice.PowerPointApi;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using PPA.Manipulation;
using PPA.Shape;
using PPA.UI.Forms;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using ALT = PPA.Core.Abstraction.Business.AlignmentType;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA.UI.Providers
{
	/// <summary>
	/// Ribbon 命令路由实现 (Refactored to use Dictionary)
	/// </summary>
	internal sealed class RibbonCommandRouter : IRibbonCommandRouter
	{
		private readonly IServiceProvider _serviceProvider;
		private readonly ILogger _logger;
		private readonly IShapeBatchHelper _shapeBatchHelper;
		private readonly Func<NETOP.Application> _getNetApp;
		private readonly Func<bool> _getTb101Press;
		private readonly Action<bool> _setTb101Press;
		private readonly ISelectionService _selectionService;
		private readonly Action<string> _invalidateControl;
		private readonly Action _invalidateRibbon;
		private readonly Dictionary<string, Action> _commandMap;

		public RibbonCommandRouter(
			IServiceProvider serviceProvider,
			ILogger logger,
			IShapeBatchHelper shapeBatchHelper,
			Func<NETOP.Application> getNetApp,
			ISelectionService selectionService,
			Func<bool> getTb101Press,
			Action<bool> setTb101Press,
			Action<string> invalidateControl,
			Action invalidateRibbon)
		{
			_serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
			_logger = logger ?? LoggerProvider.GetLogger();
			_shapeBatchHelper = shapeBatchHelper ?? throw new ArgumentNullException(nameof(shapeBatchHelper));
			_getNetApp = getNetApp ?? throw new ArgumentNullException(nameof(getNetApp));
			_selectionService = selectionService ?? throw new ArgumentNullException(nameof(selectionService));
			_getTb101Press = getTb101Press ?? throw new ArgumentNullException(nameof(getTb101Press));
			_setTb101Press = setTb101Press ?? throw new ArgumentNullException(nameof(setTb101Press));
			_invalidateControl = invalidateControl ?? throw new ArgumentNullException(nameof(invalidateControl));
			_invalidateRibbon = invalidateRibbon ?? throw new ArgumentNullException(nameof(invalidateRibbon));

			_commandMap = new Dictionary<string, Action>();
			InitializeCommands();
		}

		private void InitializeCommands()
		{
			// 1. Alignment Commands
			RegisterAlignCommand("Bt101", ALT.Left);
			RegisterAlignCommand("Bt102", ALT.Centers);
			RegisterAlignCommand("Bt103", ALT.Right);
			RegisterAlignCommand("Bt104", ALT.Horizontally);
			RegisterAlignCommand("Bt111", ALT.Top);
			RegisterAlignCommand("Bt112", ALT.Middles);
			RegisterAlignCommand("Bt113", ALT.Bottom);
			RegisterAlignCommand("Bt114", ALT.Vertically);

			// 2. Attachment Commands
			RegisterHelperCommand("Bt121", (h, a) => h.AttachLeft(a));
			RegisterHelperCommand("Bt122", (h, a) => h.AttachRight(a));
			RegisterHelperCommand("Bt123", (h, a) => h.AttachTop(a));
			RegisterHelperCommand("Bt124", (h, a) => h.AttachBottom(a));

			// 3. Size Commands
			RegisterHelperCommand("Bt201", (h, a) => h.SetEqualWidth(a));
			RegisterHelperCommand("Bt202", (h, a) => h.SetEqualHeight(a));
			RegisterHelperCommand("Bt203", (h, a) => h.SetEqualSize(a));
			RegisterHelperCommand("Bt204", (h, a) => h.SwapSize(a));
			RegisterHelperCommand("Bt211", (h, a) => h.StretchLeft(a));
			RegisterHelperCommand("Bt212", (h, a) => h.StretchRight(a));
			RegisterHelperCommand("Bt213", (h, a) => h.StretchTop(a));
			RegisterHelperCommand("Bt214", (h, a) => h.StretchBottom(a));

			// 4. Guide Commands
			RegisterHelperCommand("Bt301", (h, a) => h.GuideAlignLeft(a));
			RegisterHelperCommand("Bt302", (h, a) => h.GuideAlignHCenter(a));
			RegisterHelperCommand("Bt303", (h, a) => h.GuideAlignRight(a));
			RegisterHelperCommand("Bt311", (h, a) => h.GuideAlignTop(a));
			RegisterHelperCommand("Bt312", (h, a) => h.GuideAlignVCenter(a));
			RegisterHelperCommand("Bt313", (h, a) => h.GuideAlignBottom(a));
			RegisterHelperCommand("Bt321", (h, a) => h.GuidesStretchWidth(a));
			RegisterHelperCommand("Bt322", (h, a) => h.GuidesStretchHeight(a));
			RegisterHelperCommand("Bt323", (h, a) => h.GuidesStretchSize(a));

			// 5. Other Commands
			_commandMap["Bt401"] = () => _shapeBatchHelper.ToggleShapeVisibility(_getNetApp());
			
			_commandMap["Bt402"] = () =>
			{
				var netApp = _getNetApp();
				var validApp = ApplicationHelper.EnsureValidNetApplication(netApp);
				if (validApp != null)
				{
					MSOICrop.CropShapesToSlide(validApp);
				}
				else
				{
					_logger.LogWarning("Bt402: 无法获取有效的 Application");
				}
			};

			_commandMap["Bt501"] = () =>
			{
				var helper = ResolveTableBatchHelper();
				if (helper == null) { _logger.LogWarning("无法获取 ITableBatchHelper 服务"); return; }
				helper.FormatTables(_getNetApp());
			};

			_commandMap["Bt502"] = () =>
			{
				var helper = ResolveTextBatchHelper();
				if (helper == null) { _logger.LogWarning("无法获取 ITextBatchHelper 服务"); return; }
				helper.FormatText(_getNetApp());
			};

			_commandMap["Bt503"] = () =>
			{
				var helper = ResolveChartBatchHelper();
				if (helper == null) { _logger.LogWarning("无法获取 IChartBatchHelper 服务"); return; }
				helper.FormatCharts(_getNetApp());
			};

			_commandMap["Bt601"] = () => _shapeBatchHelper.CreateBoundingBox(_getNetApp());
		}

		private void RegisterAlignCommand(string id, ALT alignmentType)
		{
			_commandMap[id] = () =>
			{
				var netApp = _getNetApp();
				var alignHelper = ResolveAlignHelper();
				if (netApp != null && alignHelper != null)
				{
					alignHelper.ExecuteAlignment(netApp, alignmentType, _getTb101Press());
				}
			};
		}

		private void RegisterHelperCommand(string id, Action<AlignHelper, NETOP.Application> action)
		{
			_commandMap[id] = () =>
			{
				var netApp = _getNetApp();
				var alignHelper = ResolveAlignHelper();
				if (netApp != null && alignHelper != null)
				{
					action(alignHelper, netApp);
				}
			};
		}

		public bool ExecuteButtonCommand(string buttonId)
		{
			var netApp = _getNetApp();
			if (netApp == null)
			{
				_logger.LogWarning("Application 不可用，无法执行操作");
				return false;
			}

			// 在执行对齐操作前刷新切换按钮 UI
			if (buttonId.StartsWith("Bt10") || buttonId.StartsWith("Bt11"))
			{
				_invalidateControl("Tb101");
			}

			if (_commandMap.TryGetValue(buttonId, out var action))
			{
				try
				{
					action();
					return true;
				}
				catch (Exception ex)
				{
					_logger.LogError($"执行按钮命令失败 {buttonId}: {ex.Message}", ex);
					return false;
				}
			}

			_logger.LogWarning($"未知按钮ID: {buttonId}");
			return false;
		}

		public bool HandleToggleButton(Office.IRibbonControl control, bool pressed)
		{
			if (control.Id != "Tb101")
			{
				return false;
			}

			try
			{
				int shapeCount = _selectionService.GetSelectedShapeCount();
				var commandExecutor = _serviceProvider.GetService<ICommandExecutor>();

				if (shapeCount >= 2)
				{
					// 大于等于2个对象：切换状态
					bool previousState = _getTb101Press();
					_setTb101Press(pressed);

					if (previousState != pressed && commandExecutor != null)
					{
						string msoCommand = pressed
							? OfficeCommands.ObjectsAlignRelativeToContainerSmart
							: OfficeCommands.ObjectsAlignSelectedSmart;

						bool success = commandExecutor.ExecuteMso(msoCommand);
						if (success)
						{
							_logger.LogInformation($"切换状态并执行 MSO 命令 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 命令: {msoCommand}, 选中形状数: {shapeCount}");
						}
						else
						{
							_logger.LogWarning($"切换状态但 MSO 命令执行失败 | {control.Id}: {(pressed ? "幻灯片" : "所选对象")}, 命令: {msoCommand}, 选中形状数: {shapeCount}");
						}
					}
				}
				else
				{
					// 小于2个对象：设置为对齐幻灯片
					bool previousState = _getTb101Press();
					_setTb101Press(true);

					if (!previousState && commandExecutor != null)
					{
						bool success = commandExecutor.ExecuteMso(OfficeCommands.ObjectsAlignRelativeToContainerSmart);
						if (!success)
						{
							_logger.LogWarning($"设置为对齐幻灯片但 MSO 命令执行失败 | {control.Id}: 命令: {OfficeCommands.ObjectsAlignRelativeToContainerSmart}, 选中形状数: {shapeCount}");
						}
					}
				}

				_invalidateControl("Tb101");
				return true;
			}
			catch (Exception ex)
			{
				_logger.LogError($"切换按钮点击事件错误 | {control.Id}: {ex.Message}", ex);
				return false;
			}
		}

		public bool HandleMenuAction(Office.IRibbonControl control)
		{
			try
			{
				switch (control.Id)
				{
					case "MenuLang_zhCN":
						return ChangeLanguage("zh-CN", "语言已切换为中文", "Language change failed");

					case "MenuLang_enUS":
						return ChangeLanguage("en-US", "Language switched to English", "Language change failed");

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
			}
			catch (Exception ex)
			{
				_logger.LogError($"菜单项操作错误 {control.Id}: {ex.Message}", ex);
				Toast.Show($"操作失败: {ex.Message}", Toast.ToastType.Error);
				return false;
			}
		}

		private bool ChangeLanguage(string cultureCode, string successMsgKey, string errorMsgKey)
		{
			bool ok = ResourceManager.SetLanguage(cultureCode);
			if (ok)
			{
				Toast.Show(ResourceManager.GetString("Settings_LanguageChanged", successMsgKey), Toast.ToastType.Success);
				_invalidateRibbon();
			}
			else
			{
				Toast.Show(ResourceManager.GetString("Settings_LanguageChangeFailed", errorMsgKey), Toast.ToastType.Error);
			}
			return true;
		}

		#region Private Helper Methods

		private AlignHelper ResolveAlignHelper()
		{
			var service = _serviceProvider.GetService<IAlignHelper>();
			if (service is AlignHelper alignHelper)
			{
				return alignHelper;
			}
			return new AlignHelper();
		}

		private ITextBatchHelper ResolveTextBatchHelper() => _serviceProvider.GetService<ITextBatchHelper>();

		private IChartBatchHelper ResolveChartBatchHelper() => _serviceProvider.GetService<IChartBatchHelper>();

		private ITableBatchHelper ResolveTableBatchHelper() => _serviceProvider.GetService<ITableBatchHelper>();

		private void ShowSettingsDialog()
		{
			try
			{
				using var settingsForm = new SettingsForm();
				settingsForm.ShowDialog();
			}
			catch (Exception ex)
			{
				_logger.LogError($"显示设置对话框失败: {ex.Message}", ex);
				Toast.Show($"打开设置窗口失败: {ex.Message}", Toast.ToastType.Error);
			}
		}

		private void ShowAboutDialog()
		{
			try
			{
				using var aboutForm = new AboutForm();
				aboutForm.ShowDialog();
			}
			catch (Exception ex)
			{
				_logger.LogError($"显示关于对话框失败: {ex.Message}", ex);
				Toast.Show($"打开关于窗口失败: {ex.Message}", Toast.ToastType.Error);
			}
		}

		#endregion Private Helper Methods
	}
}

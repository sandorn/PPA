using System;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Utilities
{
	/// <summary>
	/// Office 原生命令执行器
	/// 提供执行 PowerPoint 内置菜单命令和功能区命令的功能
	/// </summary>
	public class CommandExecutor : ICommandExecutor
	{
		private readonly IApplication _application;

		/// <summary>
		/// 构造函数，通过依赖注入获取应用程序实例
		/// </summary>
		/// <param name="application">应用程序抽象接口</param>
		public CommandExecutor(IApplication application)
		{
			_application = application ?? throw new ArgumentNullException(nameof(application));
		}

		/// <summary>
		/// 通过 MSO 命令名称执行命令（推荐方式）
		/// </summary>
		/// <param name="msoCommandName">MSO 命令名称，例如 "Paste", "Copy", "Bold"</param>
		/// <returns>是否执行成功</returns>
		public bool ExecuteMso(string msoCommandName)
		{
			if(string.IsNullOrWhiteSpace(msoCommandName))
			{
				Profiler.LogMessage("ExecuteMso: 命令名称为空", "WARN");
				return false;
			}

			return ExHandler.Run(() =>
			{
				var nativeApp = GetNativeApplication();
				if(nativeApp == null)
				{
					Profiler.LogMessage($"ExecuteMso: 无法获取原生应用程序对象", "ERROR");
					return false;
				}

				try
				{
					Profiler.LogMessage($"ExecuteMso: 执行命令 '{msoCommandName}'", "INFO");
					nativeApp.CommandBars.ExecuteMso(msoCommandName);
					Profiler.LogMessage($"ExecuteMso: 命令 '{msoCommandName}' 执行成功", "INFO");
					return true;
				}
				catch(System.Exception ex)
				{
					Profiler.LogMessage($"ExecuteMso: 执行命令 '{msoCommandName}' 失败: {ex.Message}", "ERROR");
					return false;
				}
			}, $"执行 MSO 命令: {msoCommandName}");
		}

		/// <summary>
		/// 通过命令 ID 执行命令
		/// </summary>
		/// <param name="commandId">命令 ID</param>
		/// <returns>是否执行成功</returns>
		public bool ExecuteCommandById(int commandId)
		{
			return ExHandler.Run(() =>
			{
				var nativeApp = GetNativeApplication();
				if(nativeApp == null)
				{
					Profiler.LogMessage($"ExecuteCommandById: 无法获取原生应用程序对象", "ERROR");
					return false;
				}

				try
				{
					Profiler.LogMessage($"ExecuteCommandById: 执行命令 ID {commandId}", "INFO");
					
					// 方法1：通过 FindControl 查找并执行
					// FindControl 方法签名：FindControl(Type, Id, Tag, Visible)
					var control = nativeApp.CommandBars.FindControl(
						NetOffice.OfficeApi.Enums.MsoControlType.msoControlButton, 
						commandId, 
						Type.Missing, 
						true);
					if(control != null)
					{
						control.Execute();
						Profiler.LogMessage($"ExecuteCommandById: 命令 ID {commandId} 执行成功（通过 FindControl）", "INFO");
						return true;
					}

					// 方法2：通过 Run 方法执行
					nativeApp.Run("ExecuteCommand", new object[] { commandId });
					Profiler.LogMessage($"ExecuteCommandById: 命令 ID {commandId} 执行成功（通过 Run）", "INFO");
					return true;
				}
				catch(System.Exception ex)
				{
					Profiler.LogMessage($"ExecuteCommandById: 执行命令 ID {commandId} 失败: {ex.Message}", "ERROR");
					return false;
				}
			}, $"执行命令 ID: {commandId}");
		}

		/// <summary>
		/// 通过菜单路径执行命令（例如 "File|Save As"）
		/// </summary>
		/// <param name="menuPath">菜单路径，使用 "|" 分隔层级</param>
		/// <returns>是否执行成功</returns>
		public bool ExecuteMenuPath(string menuPath)
		{
			if(string.IsNullOrWhiteSpace(menuPath))
			{
				Profiler.LogMessage("ExecuteMenuPath: 菜单路径为空", "WARN");
				return false;
			}

			return ExHandler.Run(() =>
			{
				var nativeApp = GetNativeApplication();
				if(nativeApp == null)
				{
					Profiler.LogMessage($"ExecuteMenuPath: 无法获取原生应用程序对象", "ERROR");
					return false;
				}

				try
				{
					Profiler.LogMessage($"ExecuteMenuPath: 开始执行菜单路径 '{menuPath}'", "INFO");
					
					var parts = menuPath.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
					if(parts.Length == 0)
					{
						Profiler.LogMessage($"ExecuteMenuPath: 菜单路径格式无效", "ERROR");
						return false;
					}

					Profiler.LogMessage($"ExecuteMenuPath: 菜单路径分割为 {parts.Length} 部分: [{string.Join(", ", parts)}]", "DEBUG");

					// 获取命令栏（默认使用 Standard 工具栏）
					Profiler.LogMessage($"ExecuteMenuPath: 尝试获取命令栏 '{parts[0]}'", "DEBUG");
					NetOffice.OfficeApi.CommandBar commandBar = null;
					try
					{
						commandBar = nativeApp.CommandBars[parts[0]] as NetOffice.OfficeApi.CommandBar;
					}
					catch(Exception ex)
					{
						Profiler.LogMessage($"ExecuteMenuPath: 获取命令栏 '{parts[0]}' 失败: {ex.Message}", "DEBUG");
					}
					
					if(commandBar == null)
					{
						// 如果第一部分不是命令栏名称，尝试作为菜单栏查找
						Profiler.LogMessage($"ExecuteMenuPath: 未找到命令栏 '{parts[0]}'，尝试使用 'Menu Bar'", "DEBUG");
						try
						{
							commandBar = nativeApp.CommandBars["Menu Bar"] as NetOffice.OfficeApi.CommandBar;
						}
						catch(Exception ex)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 获取 'Menu Bar' 失败: {ex.Message}", "DEBUG");
						}
						
						if(commandBar == null)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 无法找到命令栏 '{parts[0]}' 或 'Menu Bar'", "ERROR");
							return false;
						}
					}

					Profiler.LogMessage($"ExecuteMenuPath: 成功获取命令栏 '{commandBar.Name}'，包含 {commandBar.Controls.Count} 个控件", "DEBUG");

					// 遍历菜单路径
					object currentControl = commandBar;
					for(int i = 0; i < parts.Length; i++)
					{
						var part = parts[i].Trim();
						if(string.IsNullOrEmpty(part)) continue;

						Profiler.LogMessage($"ExecuteMenuPath: 处理路径部分 [{i}]: '{part}'", "DEBUG");

						// 如果是第一部分且是命令栏名称，跳过
						if(i == 0 && currentControl == commandBar && commandBar.Name == part)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 跳过第一部分（命令栏名称）", "DEBUG");
							continue;
						}

						// 获取当前控件的子控件集合
						NetOffice.OfficeApi.CommandBarControls controls = null;
						if(currentControl is NetOffice.OfficeApi.CommandBar bar)
						{
							controls = bar.Controls;
							Profiler.LogMessage($"ExecuteMenuPath: 当前是 CommandBar，包含 {controls.Count} 个子控件", "DEBUG");
						}
						else if(currentControl is NetOffice.OfficeApi.CommandBarPopup popup)
						{
							// 只有 CommandBarPopup 类型才有 Controls 属性
							controls = popup.Controls;
							Profiler.LogMessage($"ExecuteMenuPath: 当前是 CommandBarPopup '{popup.Caption}'，包含 {controls.Count} 个子控件", "DEBUG");
						}
						else
						{
							Profiler.LogMessage($"ExecuteMenuPath: 当前控件类型: {currentControl.GetType().Name}，无法获取子控件", "DEBUG");
						}

						if(controls == null)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 无法获取控件 '{part}' 的子控件", "ERROR");
							return false;
						}

						// 调试：列出所有可用的控件
						Profiler.LogMessage($"ExecuteMenuPath: 可用控件列表:", "DEBUG");
						try
						{
							foreach(NetOffice.OfficeApi.CommandBarControl ctrl in controls)
							{
								Profiler.LogMessage($"ExecuteMenuPath:   - Caption: '{ctrl.Caption}', Tag: '{ctrl.Tag}', Type: {ctrl.Type}", "DEBUG");
							}
						}
						catch(Exception ex)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 列出控件时出错: {ex.Message}", "DEBUG");
						}

						// 查找控件
						NetOffice.OfficeApi.CommandBarControl foundControl = null;
						try
						{
							foundControl = controls[part] as NetOffice.OfficeApi.CommandBarControl;
							if(foundControl != null)
							{
								Profiler.LogMessage($"ExecuteMenuPath: 通过名称直接找到控件 '{part}'", "DEBUG");
							}
						}
						catch(Exception ex)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 通过名称查找失败: {ex.Message}，尝试遍历查找", "DEBUG");
						}

						if(foundControl == null)
						{
							// 如果通过名称查找失败，尝试遍历查找
							Profiler.LogMessage($"ExecuteMenuPath: 遍历查找控件 '{part}'", "DEBUG");
							foreach(NetOffice.OfficeApi.CommandBarControl ctrl in controls)
							{
								string caption = ctrl.Caption ?? "";
								string tag = ctrl.Tag ?? "";
								
								// 移除快捷键标记（如 (&A)）进行比较
								string captionWithoutAccel = System.Text.RegularExpressions.Regex.Replace(caption, @"\s*\(&[^)]+\)\s*", "");
								
								// 精确匹配
								if(caption == part || tag == part)
								{
									foundControl = ctrl;
									Profiler.LogMessage($"ExecuteMenuPath: 通过精确匹配找到控件: Caption='{caption}', Tag='{tag}'", "DEBUG");
									break;
								}
								// 移除快捷键标记后匹配
								else if(captionWithoutAccel == part || captionWithoutAccel.Contains(part) || part.Contains(captionWithoutAccel))
								{
									foundControl = ctrl;
									Profiler.LogMessage($"ExecuteMenuPath: 通过部分匹配找到控件: Caption='{caption}' (移除快捷键后: '{captionWithoutAccel}'), Tag='{tag}'", "DEBUG");
									break;
								}
							}
						}

						if(foundControl == null)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 无法找到控件 '{part}'", "ERROR");
							return false;
						}

						Profiler.LogMessage($"ExecuteMenuPath: 找到控件 '{foundControl.Caption}' (类型: {foundControl.Type}, Tag: '{foundControl.Tag}')", "DEBUG");

						// 如果是最后一个部分，执行命令
						if(i == parts.Length - 1)
						{
							Profiler.LogMessage($"ExecuteMenuPath: 执行最终命令 '{foundControl.Caption}'", "DEBUG");
							foundControl.Execute();
							Profiler.LogMessage($"ExecuteMenuPath: 菜单路径 '{menuPath}' 执行成功", "INFO");
							return true;
						}

						// 否则继续遍历
						currentControl = foundControl;
						Profiler.LogMessage($"ExecuteMenuPath: 继续遍历，当前控件: '{foundControl.Caption}'", "DEBUG");
					}

					Profiler.LogMessage($"ExecuteMenuPath: 菜单路径 '{menuPath}' 执行失败（未找到最终控件）", "ERROR");
					return false;
				}
				catch(System.Exception ex)
				{
					Profiler.LogMessage($"ExecuteMenuPath: 执行菜单路径 '{menuPath}' 失败: {ex.Message}", "ERROR");
					Profiler.LogMessage($"ExecuteMenuPath: 异常堆栈: {ex.StackTrace}", "DEBUG");
					return false;
				}
			}, $"执行菜单路径: {menuPath}");
		}

		/// <summary>
		/// 获取原生 PowerPoint 应用程序对象
		/// </summary>
		/// <returns>原生应用程序对象，如果无法获取则返回 null</returns>
		private NETOP.Application GetNativeApplication()
		{
			try
			{
				// 尝试从 IApplication 获取原生对象
				if(_application is IComWrapper<NETOP.Application> typed)
				{
					return typed.NativeObject;
				}

				if(_application is IComWrapper wrapper)
				{
					return wrapper.NativeObject as NETOP.Application;
				}

				// 如果无法从抽象接口获取，尝试从全局获取
				var addIn = Globals.ThisAddIn;
				if(addIn?.NetApp != null)
				{
					return addIn.NetApp;
				}

				return null;
			}
			catch(System.Exception ex)
			{
				Profiler.LogMessage($"GetNativeApplication: 获取原生应用程序对象失败: {ex.Message}", "ERROR");
				return null;
			}
		}
	}
}


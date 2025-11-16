using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PPA.Utilities
{
	/// <summary>
	/// Office 原生命令执行器
	/// 提供执行 PowerPoint 内置菜单命令和功能区命令的功能
	/// </summary>
	/// <remarks>
	/// 构造函数，通过依赖注入获取应用程序实例
	/// </remarks>
	/// <param name="application">应用程序抽象接口</param>
	public class CommandExecutor(IApplication application) : ICommandExecutor
	{
		private readonly IApplication _abstractApp = application ?? throw new ArgumentNullException(nameof(application));


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

			return ExHandler.Run<bool>(() =>
			{
				var nativeApp = ApplicationHelper.GetNetOfficeApplication(_abstractApp);
				if(nativeApp == null)
				{
					Profiler.LogMessage($"ExecuteMso: 无法获取 NetOffice Application 对象", "ERROR");
					return false;
				}

				Profiler.LogMessage($"ExecuteMso: 执行命令 '{msoCommandName}'", "INFO");
				nativeApp.CommandBars.ExecuteMso(msoCommandName);
				Profiler.LogMessage($"ExecuteMso: 命令 '{msoCommandName}' 执行成功", "INFO");
				return true;
			}, $"执行 MSO 命令: {msoCommandName}", defaultValue: false);
		}

		/// <summary>
		/// 通过命令 ID 执行命令并返回详细结果
		/// </summary>
		/// <remarks>
		/// 使用原生 COM 对象的 FindControl 方法查找命令。如果 FindControl 失败，则返回失败结果。
		/// </remarks>
		/// <param name="commandId">命令 ID</param>
		/// <returns>命令执行结果详情</returns>
		public CommandExecutionResult ExecuteCommandById(int commandId)
		{
			var result = new CommandExecutionResult { CommandId = commandId };

			return ExHandler.Run<CommandExecutionResult>(() =>
			{
				// 使用原生 COM 对象的 FindControl 方法
				MSOP.Application nativeApp = ApplicationHelper.GetNativeComApplication(_abstractApp);
				if(nativeApp == null)
				{
					result.Success = false;
					result.ErrorMessage = "无法获取原生 COM Application 对象";
					return result;
				}

				try
				{
					// 原生 COM 对象的 FindControl 方法，使用 Id 参数查找
					Office.CommandBarControl control = nativeApp.CommandBars.FindControl(Id: commandId);

					if(control == null)
					{
						result.Success = false;
						result.ErrorMessage = "未找到对应的命令控件";
						return result;
					}

					result.ControlFound = true;
					result.ControlCaption = control.Caption;
					result.ControlType = control.Type.ToString();
					
					try
					{
						result.IsEnabled = control.Enabled;
					}
					catch
					{
						result.IsEnabled = false;
					}

					if(!result.IsEnabled)
					{
						result.Success = false;
						result.ErrorMessage = "命令控件不可用";
						return result;
					}

					control.Execute();
					result.Success = true;
					result.ExecutionTime = DateTime.Now;

					Profiler.LogMessage($"ExecuteCommandById Success: ID={commandId}, Caption={control.Caption}", "INFO");
					return result;
				}
				catch(Exception ex)
				{
					result.Success = false;
					result.ErrorMessage = $"FindControl 异常: {ex.Message}";
					result.Exception = ex;

					Profiler.LogMessage($"ExecuteCommandById Error: ID={commandId}, Error={ex.Message}", "ERROR");
					return result;
				}
			}, $"执行命令详细: {commandId}", defaultValue: result);
		}


		/// <summary>
		/// 通过菜单路径执行命令（例如 "文件|另存为为"）
		/// </summary>
		/// <param name="menuPath">菜单路径，使用 "|" 分隔层级</param>
		/// <returns>是否执行成功</returns>
		public bool ExecuteMenuPath(string menuPath)
		{
			if (string.IsNullOrWhiteSpace(menuPath))
			{
				Profiler.LogMessage("ExecuteMenuPath: 菜单路径为空", "WARN");
				return false;
			}

			return ExHandler.Run(() =>
			{
				MSOP.Application nativeApp = ApplicationHelper.GetNativeComApplication(_abstractApp);
				if (nativeApp == null)
				{
					Profiler.LogMessage("ExecuteMenuPath: 无法获取原生 COM Application 对象", "ERROR");
					return false;
				}

				Profiler.LogMessage($"ExecuteMenuPath: 开始执行 '{menuPath}'", "INFO");
				
				var parts = menuPath.Split(['|'], StringSplitOptions.RemoveEmptyEntries);
				if (parts.Length == 0)
				{
					Profiler.LogMessage("ExecuteMenuPath: 菜单路径格式无效", "ERROR");
					return false;
				}

				Profiler.LogMessage($"ExecuteMenuPath: 路径分割为 {parts.Length} 部分", "DEBUG");

				// 获取命令栏
				Office.CommandBar commandBar = null;
				try
				{
					commandBar = nativeApp.CommandBars[parts[0]] as Office.CommandBar;
				}
				catch { }

				if (commandBar == null)
				{
					try
					{
						commandBar = nativeApp.CommandBars["Menu Bar"] as Office.CommandBar;
					}
					catch { }
				}

				if (commandBar == null)
				{
					Profiler.LogMessage("ExecuteMenuPath: 无法获取命令栏", "ERROR");
					return false;
				}

				Profiler.LogMessage($"ExecuteMenuPath: 使用命令栏 '{commandBar.Name}'", "DEBUG");

				// 遍历菜单路径
				object current = commandBar;
				int startIndex = current == commandBar && commandBar.Name == parts[0] ? 1 : 0;

				for (int i = startIndex; i < parts.Length; i++)
				{
					var part = parts[i].Trim();
					if (string.IsNullOrEmpty(part)) continue;

					Profiler.LogMessage($"ExecuteMenuPath: 查找 '{part}'", "DEBUG");

					// 获取子控件
					var controls = GetChildControls(current);
					if (controls == null)
					{
						Profiler.LogMessage($"ExecuteMenuPath: 无法获取子控件", "ERROR");
						return false;
					}

					// 输出所有控件信息（用于调试）
					// LogAllControls(controls);

					// 查找控件
					var control = FindControl(controls, part);
					if (control == null)
					{
						Profiler.LogMessage($"ExecuteMenuPath: 未找到控件 '{part}'", "ERROR");
						return false;
					}

					Profiler.LogMessage($"ExecuteMenuPath: 找到 '{control.Caption}'|Id:'{control.Id}'|Type:{control.Type}", "DEBUG");

					// 如果是最后一个控件，执行命令
					if (i == parts.Length - 1)
					{
						return ExecuteFinalControl(control, menuPath);
					}

					// 继续遍历
					current = control;
				}

				Profiler.LogMessage($"ExecuteMenuPath: 未找到最终控件", "ERROR");
				return false;
			}, $"执行菜单路径: {menuPath}", false);
		}

		// 辅助方法
		private Office.CommandBarControls GetChildControls(object current)
		{
			return current switch
			{
				Office.CommandBar bar => bar.Controls,
				Office.CommandBarPopup popup => popup.Controls,
				_ => null
			};
		}

		private Office.CommandBarControl FindControl(Office.CommandBarControls controls, string searchText)
		{
			if (controls == null) return null;

			// 先尝试直接通过名称查找
			try
			{
				var control = controls[searchText];
				if (control != null) return control;
			}
			catch{/*忽略异常，继续遍历查找*/}

			// 遍历查找
			string searchTextLower = searchText.ToLowerInvariant();
			int controlCount = controls.Count;
			for (int i = 1; i <= controlCount; i++)
			{
				try
				{
					var control = controls[i];
					if (control == null) continue;

					string caption = control.Caption ?? "";
					string cleanCaption = System.Text.RegularExpressions.Regex.Replace(caption, @"\s*\(&[^)]+\)\s*", "").Trim();
					string captionLower = caption.ToLowerInvariant();
					string cleanCaptionLower = cleanCaption.ToLowerInvariant();
					
					if (caption.Equals(searchText, StringComparison.OrdinalIgnoreCase) ||
						cleanCaption.Equals(searchText, StringComparison.OrdinalIgnoreCase) ||
						cleanCaptionLower.Contains(searchTextLower) ||
						searchTextLower.Contains(cleanCaptionLower))
					{
						return control;
					}
				}
				catch
				{
					// 跳过无法访问的控件
					continue;
				}
			}

			return null;
		}
		private bool ExecuteFinalControl(Office.CommandBarControl control, string menuPath)
		{
			if (!control.Enabled)
			{
				Profiler.LogMessage($"ExecuteMenuPath: 控件 '{control.Caption}' 被禁用", "WARN");
				return false;
			}

			control.Execute();
			Profiler.LogMessage($"ExecuteMenuPath: '{menuPath}' 执行成功|{control.Caption}|{control.Type}|{control.Id}", "INFO");
			return true;
		}

		/// <summary>
		/// 输出控件集合中的所有控件信息（用于调试）
		/// </summary>
		private void LogAllControls(Office.CommandBarControls controls)
		{
			if (controls == null) return;

			Profiler.LogMessage($"ExecuteMenuPath: 当前控件集合包含 {controls.Count} 个控件", "DEBUG");
			int controlCount = controls.Count;
			for (int j = 1; j <= controlCount; j++)
			{
				Office.CommandBarControl ctrl = null;
				try
				{
					ctrl = controls[j];
				}
				catch (Exception ex)
				{
					Profiler.LogMessage($"ExecuteMenuPath:   控件[{j}]: 无法访问 - {ex.Message}", "DEBUG");
					continue;
				}

				if (ctrl == null) continue;

				try
				{
					string ctrlCaption = ctrl.Caption ?? "";
					string ctrlType = ctrl.Type.ToString();
					int ctrlId = ctrl.Id;
					Profiler.LogMessage($"ExecuteMenuPath:   控件[{j}]: Caption='{ctrlCaption}', Type={ctrlType}, ID={ctrlId}", "DEBUG");
				}
				catch (Exception ex)
				{
					Profiler.LogMessage($"ExecuteMenuPath:   控件[{j}]: 获取属性失败 - {ex.Message}", "DEBUG");
				}
			}
		}
	}
}


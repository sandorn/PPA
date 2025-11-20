using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using System;
using Office = NetOffice.OfficeApi;

namespace PPA.Utilities
{
	/// <summary>
	/// Office 原生命令执行器 提供执行 PowerPoint 内置菜单命令和功能区命令的功能
	/// </summary>
	/// <remarks>
	/// 此类实现了 <see cref="ICommandExecutor" /> 接口，提供三种命令执行方式：
	/// <list type="bullet">
	/// <item>
	/// <description> <see cref="ExecuteMso(string)" /> - 通过 MSO 命令名称执行（推荐方式） </description>
	/// </item>
	/// <item>
	/// <description> <see cref="ExecuteCommandById(int)" /> - 通过命令 ID 执行（传统方式） </description>
	/// </item>
	/// <item>
	/// <description> <see cref="ExecuteMenuPath(string)" /> - 通过菜单路径执行（例如 "File|Save As"） </description>
	/// </item>
	/// </list>
	/// 优先使用 <see cref="ExecuteMso(string)" /> 方法，因为它更简洁且跨版本兼容性更好。
	/// </remarks>
	/// <remarks> 初始化 CommandExecutor 实例 </remarks>
	/// <param name="application"> 应用程序抽象接口，不能为 null </param>
	/// <param name="logger"> 日志记录器，如果为 null 则使用默认日志记录器 </param>
	/// <exception cref="ArgumentNullException"> 当 application 为 null 时抛出 </exception>
	public class CommandExecutor(IApplication application,ILogger logger = null):ICommandExecutor
	{
		private readonly IApplication _abstractApp = application??throw new ArgumentNullException(nameof(application));
		private readonly ILogger _logger = logger??LoggerProvider.GetLogger();

		/// <summary>
		/// 通过 MSO 命令名称执行命令（推荐方式）
		/// </summary>
		/// <param name="msoCommandName"> MSO 命令名称，例如 "Paste", "Copy", "Bold" </param>
		/// <returns> 是否执行成功 </returns>
		public bool ExecuteMso(string msoCommandName)
		{
			if(string.IsNullOrWhiteSpace(msoCommandName))
			{
				_logger.LogWarning("命令名称为空");
				return false;
			}

			return ExHandler.Run<bool>(() =>
			{
				var netApp = ApplicationHelper.EnsureValidNetApplication(_abstractApp);
				if(netApp==null)
				{
					_logger.LogError("无法获取有效的 Application");
					return false;
				}

				netApp.CommandBars.ExecuteMso(msoCommandName);
				_logger.LogInformation($"命令 '{msoCommandName}' 执行成功");
				return true;
			},defaultValue: false);
		}

		/// <summary>
		/// 通过命令 ID 执行命令并返回详细结果
		/// </summary>
		/// <remarks> 使用原生 COM 对象的 FindControl 方法查找命令。如果 FindControl 失败，则返回失败结果。 </remarks>
		/// <param name="commandId"> 命令 ID </param>
		/// <returns> 命令执行结果详情 </returns>
		public CommandExecutionResult ExecuteCommandById(int commandId)
		{
			var result = new CommandExecutionResult { CommandId = commandId };

			return ExHandler.Run<CommandExecutionResult>(() =>
			{
				var netApp = ApplicationHelper.EnsureValidNetApplication(_abstractApp);
				if(netApp==null)
				{
					result.Success=false;
					result.ErrorMessage="无法获取有效的 Application";
					return result;
				}

				using var commandBars = netApp.CommandBars;
				object missing = Type.Missing;
				var control = commandBars.FindControl(missing, commandId, missing, true) as Office.CommandBarControl;

				if(control==null)
				{
					result.Success=false;
					result.ErrorMessage="未找到对应的命令控件";
					return result;
				}

				result.ControlFound=true;
				result.ControlCaption=control.Caption;
				result.ControlType=control.Type.ToString();

				try
				{
					result.IsEnabled=control.Enabled;
				} catch
				{
					result.IsEnabled=false;
				}

				if(!result.IsEnabled)
				{
					result.Success=false;
					result.ErrorMessage="命令控件不可用";
					return result;
				}

				control.Execute();
				result.Success=true;
				result.ExecutionTime=DateTime.Now;

				_logger.LogInformation($"Success: ID={commandId}, Caption={control.Caption}");
				return result;
			},$"执行命令详细: {commandId}",defaultValue: result);
		}

		/// <summary>
		/// 通过菜单路径执行命令（例如 "文件|另存为为"）
		/// </summary>
		/// <param name="menuPath"> 菜单路径，使用 "|" 分隔层级 </param>
		/// <returns> 是否执行成功 </returns>
		public bool ExecuteMenuPath(string menuPath)
		{
			if(string.IsNullOrWhiteSpace(menuPath))
			{
				_logger.LogWarning("菜单路径为空");
				return false;
			}

			return ExHandler.Run(() =>
			{
				var netApp = ApplicationHelper.EnsureValidNetApplication(_abstractApp);
				if(netApp==null)
				{
					_logger.LogError("无法获取有效的 Application");
					return false;
				}

				var parts = menuPath.Split(['|'], StringSplitOptions.RemoveEmptyEntries);
				if(parts.Length==0)
				{
					_logger.LogError("菜单路径格式无效");
					return false;
				}

				using var commandBars = netApp.CommandBars;

				Office.CommandBar commandBar = null;
				try
				{
					commandBar=commandBars[parts[0]];
				} catch { }

				if(commandBar==null)
				{
					try
					{
						commandBar=commandBars["Menu Bar"];
					} catch { }
				}

				if(commandBar==null)
				{
					_logger.LogError("无法获取命令栏");
					return false;
				}

				_logger.LogDebug($"使用命令栏 '{commandBar.Name}'");

				object current = commandBar;
				int startIndex = current == commandBar && commandBar.Name == parts[0] ? 1 : 0;

				for(int i = startIndex;i<parts.Length;i++)
				{
					var part = parts[i].Trim();
					if(string.IsNullOrEmpty(part)) continue;

					_logger.LogDebug($"查找 '{part}'");

					var controls = GetChildControls(current);
					if(controls==null)
					{
						_logger.LogError("无法获取子控件");
						return false;
					}

					var control = FindControl(controls, part);
					if(control==null)
					{
						_logger.LogError($"未找到控件 '{part}'");
						return false;
					}
					_logger.LogDebug($"找到 '{control.Caption}'|Id:'{control.Id}'|Type:{control.Type}");

					if(i==parts.Length-1)
					{
						return ExecuteFinalControl(control,menuPath);
					}

					current=control;
				}

				_logger.LogError("未找到最终控件");
				return false;
			},$"执行菜单路径: {menuPath}",false);
		}

		#region Private Helper Methods

		/// <summary>
		/// 获取子控件集合
		/// </summary>
		/// <param name="current"> 当前控件对象，可以是 CommandBar 或 CommandBarPopup </param>
		/// <returns> 子控件集合，如果 current 不是支持的类型则返回 null </returns>
		private Office.CommandBarControls GetChildControls(object current)
		{
			return current switch
			{
				Office.CommandBar bar => bar.Controls,
				Office.CommandBarPopup popup => popup.Controls,
				_ => null
			};
		}

		/// <summary>
		/// 在控件集合中查找指定名称的控件
		/// </summary>
		/// <param name="controls"> 控件集合 </param>
		/// <param name="searchText"> 要查找的控件名称（不区分大小写） </param>
		/// <returns> 找到的控件，如果未找到则返回 null </returns>
		/// <remarks>
		/// 查找策略：
		/// 1. 首先尝试通过名称直接访问（controls[searchText]）
		/// 2. 如果失败，遍历所有控件，匹配 Caption 或 Tag 属性（不区分大小写）
		/// </remarks>
		private Office.CommandBarControl FindControl(Office.CommandBarControls controls,string searchText)
		{
			if(controls==null) return null;

			// 先尝试直接通过名称查找
			try
			{
				var control = controls[searchText];
				if(control!=null) return control;
			} catch {/*忽略异常，继续遍历查找*/}

			// 遍历查找
			string searchTextLower = searchText.ToLowerInvariant();
			int controlCount = controls.Count;
			for(int i = 1;i<=controlCount;i++)
			{
				try
				{
					var control = controls[i];
					if(control==null) continue;

					string caption = control.Caption ?? "";
					string cleanCaption = System.Text.RegularExpressions.Regex.Replace(caption, @"\s*\(&[^)]+\)\s*", "").Trim();
					string captionLower = caption.ToLowerInvariant();
					string cleanCaptionLower = cleanCaption.ToLowerInvariant();

					if(caption.Equals(searchText,StringComparison.OrdinalIgnoreCase)||
						cleanCaption.Equals(searchText,StringComparison.OrdinalIgnoreCase)||
						cleanCaptionLower.Contains(searchTextLower)||
						searchTextLower.Contains(cleanCaptionLower))
					{
						return control;
					}
				} catch
				{
					// 跳过无法访问的控件
					continue;
				}
			}

			return null;
		}

		/// <summary>
		/// 执行最终控件（菜单路径的最后一个控件）
		/// </summary>
		/// <param name="control"> 要执行的控件，不能为 null </param>
		/// <param name="menuPath"> 完整的菜单路径，用于日志记录 </param>
		/// <returns> 如果控件成功执行则为 true，如果控件被禁用则返回 false </returns>
		/// <remarks> 此方法会检查控件是否启用，如果启用则执行控件的 Execute 方法。 </remarks>
		private bool ExecuteFinalControl(Office.CommandBarControl control,string menuPath)
		{
			if(!control.Enabled)
			{
				_logger.LogWarning($"控件 '{control.Caption}' 被禁用");
				return false;
			}

			control.Execute();
			_logger.LogInformation($"'{menuPath}' 执行成功|{control.Caption}|{control.Type}|{control.Id}");
			return true;
		}

		/// <summary>
		/// 输出控件集合中的所有控件信息（用于调试）
		/// </summary>
		/// <param name="controls"> 要输出的控件集合，如果为 null 则不执行任何操作 </param>
		/// <remarks> 此方法会遍历所有控件并输出其 Caption、Type 和 ID 信息，用于调试菜单路径查找问题。 </remarks>
		private void LogAllControls(Office.CommandBarControls controls)
		{
			if(controls==null) return;

			_logger.LogDebug($"当前控件集合包含 {controls.Count} 个控件");
			int controlCount = controls.Count;
			for(int j = 1;j<=controlCount;j++)
			{
				Office.CommandBarControl ctrl = null;
				try
				{
					ctrl=controls[j];
				} catch(Exception ex)
				{
					_logger.LogDebug($"控件[{j}]: 无法访问 - {ex.Message}");
					continue;
				}

				if(ctrl==null) continue;

				try
				{
					string ctrlCaption = ctrl.Caption ?? "";
					string ctrlType = ctrl.Type.ToString();
					int ctrlId = ctrl.Id;
					_logger.LogDebug($"控件[{j}]: Caption='{ctrlCaption}', Type={ctrlType}, ID={ctrlId}");
				} catch(Exception ex)
				{
					_logger.LogDebug($"控件[{j}]: 获取属性失败 - {ex.Message}");
				}
			}
		}

		#endregion Private Helper Methods
	}
}

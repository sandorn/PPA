using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Formatting;
using PPA.Shape;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.UI
{
	/// <summary>
	/// 键盘快捷键助手类 提供全局快捷键注册和管理功能 使用 Windows API RegisterHotKey 实现全局快捷键
	/// </summary>
	public static class KeyboardShortcutHelper
	{
		#region Private Fields

		private static readonly Dictionary<int,Action<NETOP.Application>> _shortcuts = [];
		private static bool _initialized = false;
		private static NETOP.Application _app;
		private static MessageWindow _messageWindow;

		#endregion Private Fields

		#region Windows API

		[DllImport("user32.dll")]
		private static extern bool RegisterHotKey(IntPtr hWnd,int id,uint fsModifiers,uint vk);

		[DllImport("user32.dll")]
		private static extern bool UnregisterHotKey(IntPtr hWnd,int id);

		private const uint MOD_CONTROL = 0x0002;
		private const uint MOD_SHIFT = 0x0004;
		private const uint MOD_ALT = 0x0001;
		private const uint MOD_WIN = 0x0008;
		internal const int WM_HOTKEY = 0x0312;

		// 快捷键 ID
		private const int HOTKEY_FORMAT_TABLES = 1;

		private const int HOTKEY_FORMAT_TEXT = 2;
		private const int HOTKEY_FORMAT_CHART = 3;
		private const int HOTKEY_CREATE_BOUNDING_BOX = 4;

		#endregion Windows API

		#region Public Methods

		/// <summary>
		/// 初始化快捷键系统
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public static void Initialize(NETOP.Application app)
		{
			if(_initialized||app==null) return;

			try
			{
				_app=app;

				// 创建消息窗口用于接收热键消息
				_messageWindow=new MessageWindow();
				_messageWindow.HotKeyPressed+=OnHotKeyPressed;

				// 注册全局快捷键
				RegisterGlobalShortcuts();

				_initialized=true;
				Profiler.LogMessage("快捷键系统初始化成功（全局快捷键）","INFO");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"快捷键系统初始化失败: {ex.Message}","WARN");
			}
		}

		/// <summary>
		/// 注销快捷键系统
		/// </summary>
		public static void Uninitialize()
		{
			if(!_initialized) return;

			try
			{
				// 注销全局快捷键
				UnregisterGlobalShortcuts();

				// 释放消息窗口
				_messageWindow?.Dispose();
				_messageWindow=null;

				_shortcuts.Clear();
				_app=null;
				_initialized=false;
				Profiler.LogMessage("快捷键系统已注销","INFO");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"快捷键系统注销失败: {ex.Message}","WARN");
			}
		}

		/// <summary>
		/// 重新注册快捷键（用于配置更新后重新加载）
		/// </summary>
		public static void ReloadShortcuts()
		{
			if(!_initialized||_app==null) return;

			try
			{
				// 先注销所有快捷键
				UnregisterGlobalShortcuts();

				// 重新注册快捷键
				RegisterGlobalShortcuts();

				Profiler.LogMessage("快捷键已重新加载","INFO");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"重新加载快捷键失败: {ex.Message}","WARN");
			}
		}

		#endregion Public Methods

		#region Private Methods

		/// <summary>
		/// 注册全局快捷键
		/// </summary>
		private static void RegisterGlobalShortcuts()
		{
			if(_messageWindow?.Handle==null) return;

			IntPtr hWnd = _messageWindow.Handle;

			// 从配置文件读取快捷键设置
			var config = FormattingConfig.Instance;
			var shortcuts = config?.Shortcuts;

			if(shortcuts==null) return;

			// 注册美化表格快捷键
			if(!string.IsNullOrWhiteSpace(shortcuts.FormatTables))
			{
				RegisterShortcut(hWnd,HOTKEY_FORMAT_TABLES,shortcuts.FormatTables,
					(app) =>
					{
						var helper = ResolveTableBatchHelper();
						if(helper==null)
						{
							Profiler.LogMessage("警告：无法获取 ITableBatchHelper 服务", "WARN");
							return;
						}
						helper.FormatTables(app);
					},"美化表格");
			}

			// 注册美化文本快捷键
			if(!string.IsNullOrWhiteSpace(shortcuts.FormatText))
			{
				RegisterShortcut(hWnd,HOTKEY_FORMAT_TEXT,shortcuts.FormatText,
					(app) =>
					{
						var helper = ResolveTextBatchHelper();
						if(helper==null)
						{
							Profiler.LogMessage("警告：无法获取 ITextBatchHelper 服务", "WARN");
							return;
						}
						helper.FormatText(app);
					},"美化文本");
			}

			// 注册美化图表快捷键
			if(!string.IsNullOrWhiteSpace(shortcuts.FormatChart))
			{
				RegisterShortcut(hWnd,HOTKEY_FORMAT_CHART,shortcuts.FormatChart,
					(app) =>
					{
						var helper = ResolveChartBatchHelper();
						if(helper==null)
						{
							Profiler.LogMessage("警告：无法获取 IChartBatchHelper 服务", "WARN");
							return;
						}
						helper.FormatCharts(app);
					},"美化图表");
			}

			// 注册插入形状快捷键
			if(!string.IsNullOrWhiteSpace(shortcuts.CreateBoundingBox))
			{
				RegisterShortcut(hWnd,HOTKEY_CREATE_BOUNDING_BOX,shortcuts.CreateBoundingBox,
					(app) => ShapeBatchHelper.Bt601_Click(app),"插入形状");
			}
		}

		private static ITextBatchHelper ResolveTextBatchHelper()
		{
			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			if(serviceProvider==null)
			{
				return null;
			}

			var batchHelper = serviceProvider.GetService<ITextBatchHelper>();
			if(batchHelper!=null)
			{
				return batchHelper;
			}

			var textHelper = serviceProvider.GetService<ITextFormatHelper>();
		var shapeHelper = ResolveShapeHelper(serviceProvider);
		if(textHelper!=null)
			{
				return new TextBatchHelper(textHelper,shapeHelper);
			}

			return null;
		}

		private static IChartBatchHelper ResolveChartBatchHelper()
		{
			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			if(serviceProvider==null)
			{
				return null;
			}

			var batchHelper = serviceProvider.GetService<IChartBatchHelper>();
			if(batchHelper!=null)
			{
				return batchHelper;
			}

			var formatHelper = serviceProvider.GetService<IChartFormatHelper>();
			var shapeHelper = ResolveShapeHelper(serviceProvider);
		if(formatHelper!=null)
			{
				return new ChartBatchHelper(formatHelper,shapeHelper);
			}

			return null;
		}

		private static ITableBatchHelper ResolveTableBatchHelper()
		{
			var serviceProvider = Globals.ThisAddIn?.ServiceProvider;
			if(serviceProvider==null)
			{
				return null;
			}

			var batchHelper = serviceProvider.GetService<ITableBatchHelper>();
			if(batchHelper!=null)
			{
				return batchHelper;
			}

			var tableHelper = serviceProvider.GetService<ITableFormatHelper>();
			var shapeHelper = ResolveShapeHelper(serviceProvider);
			if(tableHelper!=null)
			{
				return new TableBatchHelper(tableHelper,shapeHelper);
			}

			return null;
		}

		private static IShapeHelper ResolveShapeHelper(IServiceProvider serviceProvider = null)
		{
			serviceProvider ??= Globals.ThisAddIn?.ServiceProvider;
			var helper = serviceProvider?.GetService<IShapeHelper>();
			return helper ?? ShapeUtils.Default;
		}

		/// <summary>
		/// 注册单个快捷键
		/// </summary>
		private static void RegisterShortcut(IntPtr hWnd,int hotkeyId,string shortcut,
			Action<NETOP.Application> action,string actionName)
		{
			if(TryParseShortcut(shortcut,out uint modifiers,out uint vk))
			{
				if(RegisterHotKey(hWnd,hotkeyId,modifiers,vk))
				{
					_shortcuts[hotkeyId]=action;
					// 显示完整的快捷键（Ctrl + 配置的值）
					string fullShortcut = $"Ctrl+{shortcut}";
					Profiler.LogMessage($"注册快捷键: {fullShortcut} ({actionName})","INFO");
				} else
				{
					string fullShortcut = $"Ctrl+{shortcut}";
					Profiler.LogMessage($"注册快捷键 {fullShortcut} ({actionName}) 失败（可能已被占用）","WARN");
				}
			} else
			{
				Profiler.LogMessage($"快捷键格式无效: {shortcut} ({actionName})，跳过注册","WARN");
			}
		}

		/// <summary>
		/// 解析快捷键字符串 支持的格式：只配置数字或字母（如 "3", "C", "F1"），系统会自动添加 Ctrl 修饰键
		/// </summary>
		/// <param name="shortcut"> 快捷键字符串（只包含数字或字母，不包含修饰键） </param>
		/// <param name="modifiers"> 修饰键标志（始终包含 MOD_CONTROL） </param>
		/// <param name="vk"> 虚拟键码 </param>
		/// <returns> 是否解析成功 </returns>
		private static bool TryParseShortcut(string shortcut,out uint modifiers,out uint vk)
		{
			modifiers=MOD_CONTROL; // 统一使用 Ctrl 修饰键
			vk=0;

			if(string.IsNullOrWhiteSpace(shortcut))
				return false;

			// 统一转换为大写
			string key = shortcut.Trim().ToUpper();

			// 数字键 0-9
			if(key.Length==1&&char.IsDigit(key[0]))
			{
				vk=(uint) (key[0]-'0'+0x30); // '0' = 0x30, '1' = 0x31, etc.
				return true;
			}

			// 字母键 A-Z
			if(key.Length==1&&char.IsLetter(key[0]))
			{
				vk=(uint) key[0];
				return true;
			}

			// 功能键 F1-F12
			if(key.StartsWith("F")&&key.Length>=2)
			{
				if(int.TryParse(key.Substring(1),out int fNum)&&fNum>=1&&fNum<=12)
				{
					vk=(uint) (0x70+fNum-1); // F1 = 0x70, F2 = 0x71, etc.
					return true;
				}
			}

			return false;
		}

		/// <summary>
		/// 注销全局快捷键
		/// </summary>
		private static void UnregisterGlobalShortcuts()
		{
			if(_messageWindow?.Handle==null) return;

			IntPtr hWnd = _messageWindow.Handle;

			// 注销所有已注册的快捷键
			foreach(var hotkeyId in _shortcuts.Keys)
			{
				UnregisterHotKey(hWnd,hotkeyId);
			}
		}

		/// <summary>
		/// 处理热键按下事件
		/// </summary>
		private static void OnHotKeyPressed(object sender,HotKeyEventArgs e)
		{
			if(_shortcuts.TryGetValue(e.HotKeyId,out var action)&&_app!=null)
			{
				try
				{
					action(_app);
				} catch(Exception ex)
				{
					Profiler.LogMessage($"执行快捷键失败: {ex.Message}","ERROR");
				}
			}
		}

		#endregion Private Methods

		#region Shortcut Definitions

		/// <summary>
		/// 预定义的快捷键常量
		/// </summary>
		public static class Shortcuts
		{
			/// <summary>
			/// 美化图表：Ctrl+3
			/// </summary>
			public const string FormatChart = "Ctrl+3";
		}

		#endregion Shortcut Definitions
	}

	#region Message Window for HotKey

	/// <summary>
	/// 用于接收热键消息的隐藏窗口
	/// </summary>
	internal class MessageWindow:NativeWindow, IDisposable
	{
		public event EventHandler<HotKeyEventArgs> HotKeyPressed;

		public MessageWindow()
		{
			CreateHandle(new CreateParams());
		}

		protected override void WndProc(ref Message m)
		{
			if(m.Msg==KeyboardShortcutHelper.WM_HOTKEY)
			{
				int hotKeyId = m.WParam.ToInt32();
				HotKeyPressed?.Invoke(this,new HotKeyEventArgs(hotKeyId));
			}

			base.WndProc(ref m);
		}

		public void Dispose()
		{
			DestroyHandle();
		}
	}

	/// <summary>
	/// 热键事件参数
	/// </summary>
	internal class HotKeyEventArgs(int hotKeyId):EventArgs
	{
		public int HotKeyId { get; } = hotKeyId;
	}

	#endregion Message Window for HotKey
}

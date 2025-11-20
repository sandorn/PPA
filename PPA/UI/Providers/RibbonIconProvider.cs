using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using System;
using System.Collections.Generic;
using System.Drawing;
using Office = Microsoft.Office.Core;

namespace PPA.UI.Providers
{
	/// <summary>
	/// Ribbon 图标资源提供者实现
	/// </summary>
	internal sealed class RibbonIconProvider:IRibbonIconProvider
	{
		private readonly Dictionary<string,Bitmap> _iconCache;
		private readonly ILogger _logger;

		public RibbonIconProvider(ILogger logger = null)
		{
			_logger=logger??LoggerProvider.GetLogger();
			_iconCache=new Dictionary<string,Bitmap>();
		}

		/// <summary>
		/// 获取指定控件的图标
		/// </summary>
		public Bitmap GetIcon(Office.IRibbonControl control,bool? pressed = null)
		{
			try
			{
				string itemId = control.Id;

				// 处理切换按钮的特殊逻辑
				if(control.Id=="Tb101")
				{
					bool isPressed = pressed ?? false;
					itemId=isPressed ? "Tb101_1" : "Tb101_0";
				}

				if(_iconCache.TryGetValue(itemId,out Bitmap bmp))
				{
					return bmp;
				}

				_logger.LogWarning($"未找到图标: {itemId}");
				return null;
			} catch(Exception ex)
			{
				_logger.LogError($"获取图标错误 | {control.Id}: {ex.Message}",ex);
				return null;
			}
		}

		/// <summary>
		/// 预加载所有 Ribbon 图标到缓存中
		/// </summary>
		public void PreloadIcons()
		{
			if(_iconCache.Count>0) return;

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

				_logger.LogInformation($"已预加载 {_iconCache.Count} 个图标");
			} catch(Exception ex)
			{
				_logger.LogError($"预加载图标错误: {ex.Message}",ex);
			}
		}

		/// <summary>
		/// 释放所有缓存的图标资源
		/// </summary>
		public void DisposeIcons()
		{
			foreach(var kvp in _iconCache)
			{
				try
				{
					kvp.Value?.Dispose();
				} catch(Exception ex)
				{
					_logger.LogWarning($"释放图标资源时出错 | {kvp.Key}: {ex.Message}");
				}
			}
			_iconCache.Clear();
		}
	}
}

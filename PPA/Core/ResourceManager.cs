using System;
using System.Globalization;
using System.Threading;

namespace PPA.Core
{
	/// <summary>
	/// 多语言资源管理器 提供本地化字符串的加载和获取功能
	/// </summary>
	public static class ResourceManager
	{
		#region Private Fields

		private static System.Resources.ResourceManager _resourceManager;
		private static CultureInfo _currentCulture;

		#endregion Private Fields

		#region Public Properties

		/// <summary>
		/// 当前语言文化信息
		/// </summary>
		public static CultureInfo CurrentCulture
		{
			get => _currentCulture??CultureInfo.CurrentUICulture;
			set
			{
				_currentCulture=value;
				Thread.CurrentThread.CurrentUICulture=value;
			}
		}

		/// <summary>
		/// 支持的语言列表
		/// </summary>
		public static readonly string[] SupportedLanguages = { "zh-CN","en-US" };

		#endregion Public Properties

		#region Public Methods

		/// <summary>
		/// 初始化资源管理器
		/// </summary>
		/// <param name="baseName"> 资源文件的基础名称（不含语言后缀） </param>
		/// <param name="assembly"> 包含资源文件的程序集 </param>
		public static void Initialize(string baseName,System.Reflection.Assembly assembly)
		{
			_resourceManager=new System.Resources.ResourceManager(baseName,assembly);

			// 默认使用系统语言，如果系统语言不支持则使用中文
			var systemCulture = CultureInfo.CurrentUICulture;
			if(Array.IndexOf(SupportedLanguages,systemCulture.Name)>=0)
			{
				CurrentCulture=systemCulture;
			} else
			{
				CurrentCulture=new CultureInfo("zh-CN");
			}

			Profiler.LogMessage($"资源管理器初始化成功，当前语言: {CurrentCulture.Name}","INFO");
		}

		/// <summary>
		/// 获取本地化字符串
		/// </summary>
		/// <param name="key"> 资源键名 </param>
		/// <param name="defaultValue"> 默认值（如果找不到资源时使用） </param>
		/// <returns> 本地化字符串 </returns>
		public static string GetString(string key,string defaultValue = null)
		{
			if(_resourceManager==null)
			{
				Profiler.LogMessage($"资源管理器未初始化，返回默认值: {key}","WARN");
				return defaultValue??key;
			}

			try
			{
				// 只从当前语言获取，不使用后备语言 这样可以确保在英文环境下不会回退到中文
				string value = _resourceManager.GetString(key, CurrentCulture);
				return value??defaultValue??key;
			} catch(Exception ex)
			{
				Profiler.LogMessage($"获取资源字符串失败: {key}, 错误: {ex.Message}","WARN");
				return defaultValue??key;
			}
		}

		/// <summary>
		/// 获取格式化字符串（支持参数替换）
		/// </summary>
		/// <param name="key"> 资源键名 </param>
		/// <param name="args">
		/// 格式化参数。如果第一个参数是字符串且包含占位符（如 {0}），则：
		/// - 如果资源文件中找到了 key，第一个参数会被忽略（作为 defaultValue），使用资源文件中的格式和剩余参数
		/// - 如果资源文件中找不到 key，第一个参数会被用作默认格式字符串
		/// </param>
		/// <returns> 格式化后的本地化字符串 </returns>
		public static string GetString(string key,params object[] args)
		{
			// 如果没有参数，使用简单的 GetString 方法
			if(args==null||args.Length==0)
			{
				return GetString(key);
			}

			// 先尝试从资源文件获取格式字符串（不传 defaultValue，这样如果找不到会返回 key）
			string format;
			if(_resourceManager==null)
			{
				format=key;
			} else
			{
				try
				{
					// 只从当前语言获取，不使用后备语言 这样可以确保在英文环境下不会回退到中文
					format=_resourceManager.GetString(key,CurrentCulture);
					format=format??key;
				} catch(Exception ex)
				{
					Profiler.LogMessage($"[ResourceManager] 获取资源失败: {key}, 错误: {ex.Message}","WARN");
					format=key;
				}
			}

			// 判断是否在资源文件中找到：如果返回的是 key 本身，说明没找到
			bool foundInResources = format != key;

			// 检查第一个参数是否是包含占位符的字符串（可能是 defaultValue）
			string firstArg = null;
			bool firstArgIsFormatString = args.Length > 0 && args[0] is string arg && (firstArg = arg).Contains("{");

			if(firstArgIsFormatString)
			{
				if(!foundInResources)
				{
					// 资源文件中找不到，使用第一个参数作为默认格式
					format=firstArg;
					// 移除第一个参数，因为它已经被用作格式字符串
					object[] formatArgs = new object[args.Length - 1];
					Array.Copy(args,1,formatArgs,0,formatArgs.Length);
					args=formatArgs;
				} else
				{
					// 资源文件中找到了，第一个参数是 defaultValue，忽略它，使用资源文件中的格式和剩余参数
					object[] formatArgs = new object[args.Length - 1];
					Array.Copy(args,1,formatArgs,0,formatArgs.Length);
					args=formatArgs;
				}
			}
			// 如果第一个参数不是格式字符串，直接使用所有参数进行格式化

			try
			{
				return string.Format(format,args);
			} catch(Exception ex)
			{
				Profiler.LogMessage($"格式化字符串失败: {key}, 格式: {format}, 参数数量: {args?.Length??0}, 错误: {ex.Message}","WARN");
				return format;
			}
		}

		/// <summary>
		/// 切换语言
		/// </summary>
		/// <param name="cultureName"> 语言文化名称（如 "zh-CN", "en-US"） </param>
		public static void SetLanguage(string cultureName)
		{
			if(Array.IndexOf(SupportedLanguages,cultureName)<0)
			{
				Profiler.LogMessage($"不支持的语言: {cultureName}，使用默认语言","WARN");
				return;
			}

			try
			{
				CurrentCulture=new CultureInfo(cultureName);
				Profiler.LogMessage($"语言已切换为: {cultureName}","INFO");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"切换语言失败: {ex.Message}","WARN");
			}
		}

		#endregion Public Methods
	}
}

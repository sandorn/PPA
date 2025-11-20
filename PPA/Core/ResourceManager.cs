using System;
using System.Globalization;
using System.Threading;

namespace PPA.Core
{
	/// <summary>
	/// 多语言资源管理器 - 提供本地化字符串的加载和获取功能
	/// </summary>
	public static class ResourceManager
	{
		#region Private Fields

		private static System.Resources.ResourceManager _resourceManager;
		private static CultureInfo _currentCulture;
		private static readonly object _lockObject = new();

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
				lock(_lockObject)
				{
					_currentCulture=value??throw new ArgumentNullException(nameof(value));
					Thread.CurrentThread.CurrentUICulture=value;
					Thread.CurrentThread.CurrentCulture=value; // 同时设置 CurrentCulture
				}

				Profiler.LogMessage($"语言文化已设置为: {value.Name}","INFO");
			}
		}

		/// <summary>
		/// 支持的语言列表
		/// </summary>
		public static readonly string[] SupportedLanguages = ["zh-CN","en-US"];

		/// <summary>
		/// 获取默认语言（中文）
		/// </summary>
		public static CultureInfo DefaultCulture => new("zh-CN");

		#endregion Public Properties

		#region Public Methods

		/// <summary>
		/// 初始化资源管理器
		/// </summary>
		/// <param name="baseName"> 资源文件的基础名称（不含语言后缀） </param>
		/// <param name="assembly"> 包含资源文件的程序集 </param>
		public static void Initialize(string baseName,System.Reflection.Assembly assembly)
		{
			if(string.IsNullOrEmpty(baseName))
				throw new ArgumentException("资源文件基础名称不能为空",nameof(baseName));
			if(assembly==null)
				throw new ArgumentNullException(nameof(assembly));

			lock(_lockObject)
			{
				_resourceManager=new System.Resources.ResourceManager(baseName,assembly);

				// 默认使用系统语言，如果系统语言不支持则使用中文
				var systemCulture = CultureInfo.CurrentUICulture;
				if(Array.IndexOf(SupportedLanguages,systemCulture.Name)>=0)
				{
					CurrentCulture=systemCulture;
				} else
				{
					// 尝试匹配父文化（如 en 匹配 en-US）
					var parentCulture = GetSupportedParentCulture(systemCulture);
					CurrentCulture=parentCulture??DefaultCulture;
				}
			}

			Profiler.LogMessage($"资源管理器初始化成功，当前语言: {CurrentCulture.Name}","INFO");
		}

		/// <summary>
		/// 获取本地化字符串
		/// </summary>
		public static string GetString(string key,string defaultValue = null)
		{
			if(string.IsNullOrEmpty(key))
				return defaultValue??string.Empty;

			if(_resourceManager==null)
			{
				Profiler.LogMessage($"资源管理器未初始化，返回默认值: {key}","WARN");
				return defaultValue??key;
			}

			try
			{
				// 只从当前语言获取，不使用后备语言
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
		public static string GetString(string key,params object[] args)
		{
			// 如果没有参数，使用简单的 GetString 方法
			if(args==null||args.Length==0)
			{
				return GetString(key);
			}

			string format;
			bool foundInResources;

			if(_resourceManager==null)
			{
				format=key;
				foundInResources=false;
			} else
			{
				try
				{
					format=_resourceManager.GetString(key,CurrentCulture);
					foundInResources=format!=null;
					format??=key;
				} catch(Exception ex)
				{
					Profiler.LogMessage($"获取资源失败: {key}, 错误: {ex.Message}","WARN");
					format=key;
					foundInResources=false;
				}
			}

			// 处理第一个参数可能是默认格式字符串的情况
			object[] formatArgs = ProcessFormatArguments(ref format, foundInResources, args);

			try
			{
				return string.Format(format,formatArgs);
			} catch(FormatException ex)
			{
				Profiler.LogMessage($"格式化字符串失败: {key}, 格式: {format}, 参数数量: {formatArgs.Length}, 错误: {ex.Message}","WARN");
				// 返回未格式化的字符串，避免显示异常信息给用户
				return format;
			}
		}

		/// <summary>
		/// 切换语言
		/// </summary>
		public static bool SetLanguage(string cultureName)
		{
			if(string.IsNullOrEmpty(cultureName))
			{
				Profiler.LogMessage("语言名称不能为空","WARN");
				return false;
			}

			if(Array.IndexOf(SupportedLanguages,cultureName)<0)
			{
				Profiler.LogMessage($"不支持的语言: {cultureName}，使用默认语言","WARN");
				return false;
			}

			try
			{
				CurrentCulture=new CultureInfo(cultureName);
				return true;
			} catch(CultureNotFoundException ex)
			{
				Profiler.LogMessage($"语言文化不存在: {cultureName}, 错误: {ex.Message}","WARN");
				return false;
			} catch(Exception ex)
			{
				Profiler.LogMessage($"切换语言失败: {ex.Message}","WARN");
				return false;
			}
		}

		#endregion Public Methods

		#region Private Methods

		/// <summary>
		/// 处理格式化参数（可能更新 format）
		/// </summary>
		private static object[] ProcessFormatArguments(ref string format,bool foundInResources,object[] args)
		{
			string firstArg = (args != null && args.Length > 0) ? args[0] as string : null;
			bool firstArgIsFormatString = !string.IsNullOrEmpty(firstArg) && ContainsPlaceholders(firstArg);

			if(firstArgIsFormatString)
			{
				if(!foundInResources)
				{
					// 资源文件中找不到，使用第一个参数作为默认格式
					format=firstArg;
					object[] newArgs = new object[args.Length - 1];
					Array.Copy(args,1,newArgs,0,newArgs.Length);
					return newArgs;
				} else
				{
					// 资源文件中找到了，第一个参数是 defaultValue，忽略它
					object[] newArgs = new object[args.Length - 1];
					Array.Copy(args,1,newArgs,0,newArgs.Length);
					return newArgs;
				}
			}

			return args??Array.Empty<object>();
		}

		/// <summary>
		/// 检查字符串是否包含占位符
		/// </summary>
		private static bool ContainsPlaceholders(string text)
		{
			return text!=null&&text.Contains("{")&&text.Contains("}");
		}

		/// <summary>
		/// 获取支持的父级文化
		/// </summary>
		private static CultureInfo GetSupportedParentCulture(CultureInfo culture)
		{
			try
			{
				var parent = culture.Parent;
				while(parent!=null&&!parent.Equals(CultureInfo.InvariantCulture))
				{
					if(Array.IndexOf(SupportedLanguages,parent.Name)>=0)
					{
						return parent;
					}
					parent=parent.Parent;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"获取父级文化失败: {ex.Message}","DEBUG");
			}

			return null;
		}

		#endregion Private Methods
	}
}

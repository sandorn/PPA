using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using PPA.Properties;
using System;
using System.IO;
using System.Reflection;

namespace PPA.UI.Providers
{
	/// <summary>
	/// 从嵌入式资源或后备资源加载 Ribbon XML 的提供者
	/// </summary>
	internal sealed class EmbeddedRibbonXmlProvider:IRibbonXmlProvider
	{
		private readonly ILogger _logger;

		public EmbeddedRibbonXmlProvider(ILogger logger = null)
		{
			_logger=logger??LoggerProvider.GetLogger();
		}

		/// <summary>
		/// 获取 Ribbon XML 字符串 优先从嵌入式资源加载，失败则使用后备资源
		/// </summary>
		public string GetRibbonXml(string ribbonID)
		{
			try
			{
				string ribbonXml = LoadRibbonXmlFromFile();
				if(!string.IsNullOrEmpty(ribbonXml))
				{
					return ribbonXml;
				}
			} catch(Exception ex)
			{
				_logger.LogError($"从文件加载XML失败: {ex.Message}",ex);
			}

			// 后备方案：使用嵌入的资源字符串
			return Resources.RibbonXml;
		}

		/// <summary>
		/// 从嵌入式资源中加载 Ribbon XML
		/// </summary>
		/// <returns> 加载的XML字符串，如未找到则返回null </returns>
		private string LoadRibbonXmlFromFile()
		{
			try
			{
				// 从嵌入式资源加载 Ribbon.xml 资源名称格式：命名空间.文件夹.文件名
				string resourceName = "PPA.UI.Ribbon.xml";
				var assembly = Assembly.GetExecutingAssembly();

				using(var stream = assembly.GetManifestResourceStream(resourceName))
				{
					if(stream!=null)
					{
						using var reader = new StreamReader(stream);
						string xmlContent = reader.ReadToEnd();
						_logger.LogInformation("成功从嵌入式资源加载 Ribbon.xml");
						return xmlContent;
					}
				}

				_logger.LogWarning($"未找到嵌入式资源: {resourceName}，使用后备资源");
			} catch(Exception ex)
			{
				_logger.LogError($"从嵌入式资源加载 Ribbon.xml 失败: {ex.Message}",ex);
			}

			return null;
		}
	}
}

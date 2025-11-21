using NetOffice.OfficeApi.Enums;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using NETOP = NetOffice.PowerPointApi;
using PPA.Manipulation.Config;

namespace PPA.Manipulation
{
	/// <summary>
	/// PPA 配置类 用于管理表格、文本、图表的格式化样式配置和快捷键配置
	/// </summary>
	[XmlRoot("PPAConfig")]
	public class FormattingConfig:PPA.Core.Abstraction.Business.IFormattingConfig
	{
		private static ILogger Logger => LoggerProvider.GetLogger();

		#region Singleton

		private static FormattingConfig _instance;
		private static readonly object _lock = new();
		private static string _configFilePath;

		/// <summary>
		/// 获取配置实例（单例模式）
		/// </summary>
		public static FormattingConfig Instance
		{
			get
			{
				if(_instance==null)
				{
					lock(_lock)
					{
						_instance??=LoadConfig();
					}
				}
				return _instance;
			}
		}

		/// <summary>
		/// 获取配置文件路径
		/// </summary>
		private static string GetConfigFilePath()
		{
			if(_configFilePath==null)
			{
				// 使用 AppData 目录存放配置文件
				string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
				if(string.IsNullOrEmpty(appDataDir))
				{
					// 如果获取 AppData 失败，尝试使用用户目录
					appDataDir=Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)??
								 Environment.GetEnvironmentVariable("USERPROFILE")??
								 Environment.GetEnvironmentVariable("HOME")??
								 ".";
				}

				// 创建 PPA 子目录（如果不存在）
				string ppaConfigDir = Path.Combine(appDataDir, "PPA");
				if(!Directory.Exists(ppaConfigDir))
				{
					try
					{
						Directory.CreateDirectory(ppaConfigDir);
					} catch(Exception ex)
					{
						Logger.LogWarning($"创建配置目录失败: {ex.Message}，使用用户目录");
						ppaConfigDir=appDataDir;
					}
				}

				_configFilePath=Path.Combine(ppaConfigDir,"PPAConfig.xml");
			}
			return _configFilePath;
		}

		/// <summary>
		/// 加载配置文件
		/// </summary>
		private static FormattingConfig LoadConfig()
		{
			string configPath = GetConfigFilePath();

			try
			{
				// 如果配置文件存在，直接加载
				if(File.Exists(configPath))
				{
					var config = LoadConfigFromFile(configPath);
					if(config!=null)
					{
						return config;
					}
				}
			} catch(Exception ex)
			{
				Logger.LogError($"加载配置文件失败: {ex.Message}，使用默认配置",ex);
			}

			// 如果加载失败或文件不存在，返回默认配置
			var defaultConfig = new FormattingConfig();
			// 设置默认快捷键（仅在创建新配置文件时） 只配置数字或字母，系统会自动添加 Ctrl 修饰键
			defaultConfig.Shortcuts.FormatChart="3";
			defaultConfig.Save(); // 保存默认配置到文件
			return defaultConfig;
		}

		/// <summary>
		/// 从文件加载配置
		/// </summary>
		private static FormattingConfig LoadConfigFromFile(string filePath)
		{
			try
			{
				var serializer = new XmlSerializer(typeof(FormattingConfig));
				using var reader = new StreamReader(filePath, Encoding.UTF8);
				var config = (FormattingConfig)serializer.Deserialize(reader);
				Logger.LogInformation($"已加载配置文件: {filePath}");
				return config;
			} catch(Exception ex)
			{
				Logger.LogError($"从文件加载配置失败: {ex.Message}",ex);
				return null;
			}
		}

		/// <summary>
		/// 保存配置到文件
		/// </summary>
		public void Save()
		{
			try
			{
				string configPath = GetConfigFilePath();
				var serializer = new XmlSerializer(typeof(FormattingConfig));
				var ns = new XmlSerializerNamespaces();
				ns.Add("",""); // 移除命名空间

				// 先序列化到内存
				string xmlContent;
				using(var stringWriter = new StringWriterWithEncoding(Encoding.UTF8))
				{
					using(var xmlWriter = XmlWriter.Create(stringWriter,new XmlWriterSettings
					{
						Indent=true,
						IndentChars="\t",
						NewLineChars="\n",
						Encoding=Encoding.UTF8,
						OmitXmlDeclaration=false
					}))
					{
						serializer.Serialize(xmlWriter,this,ns);
					}
					xmlContent=stringWriter.ToString();
				}

				// 格式化 XML：每个属性换行
				xmlContent=FormatXmlWithAttributesOnNewLines(xmlContent);

				// 写入文件
				File.WriteAllText(configPath,xmlContent,Encoding.UTF8);

				Logger.LogInformation($"配置文件已保存: {configPath}");
			} catch(Exception ex)
			{
				Logger.LogError($"保存配置文件失败: {ex.Message}",ex);
			}
		}

		/// <summary>
		/// 格式化 XML，使每个属性换行显示
		/// </summary>
		private static string FormatXmlWithAttributesOnNewLines(string xml)
		{
			try
			{
				var lines = xml.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
				var result = new StringBuilder();

				foreach(var line in lines)
				{
					var trimmedLine = line.Trim();
					if(string.IsNullOrEmpty(trimmedLine))
						continue;

					// 计算当前行的缩进（基于制表符）
					var lineIndent = line.TakeWhile(c => c == '\t').Count();
					var indentStr = new string('\t', lineIndent);
					var attrIndentStr = new string('\t', lineIndent + 1);

					// 检查是否是开始标签或自闭合标签
					if(trimmedLine.StartsWith("<")&&trimmedLine.Contains(" ")&&!trimmedLine.StartsWith("</")&&!trimmedLine.StartsWith("<?")&&!trimmedLine.StartsWith("<!--"))
					{
						// 提取标签名和属性
						var tagMatch = Regex.Match(trimmedLine, @"<(\w+)([^>]*?)(/?>)");
						if(tagMatch.Success)
						{
							var tagName = tagMatch.Groups[1].Value;
							var attributesStr = tagMatch.Groups[2].Value.Trim();
							var closing = tagMatch.Groups[3].Value;

							if(!string.IsNullOrEmpty(attributesStr))
							{
								// 提取所有属性
								var attributes = new List<string>();
								var attrPattern = @"(\S+)\s*=\s*""([^""]*)""";
								var attrMatches = Regex.Matches(attributesStr, attrPattern);

								foreach(Match attrMatch in attrMatches)
								{
									var attrName = attrMatch.Groups[1].Value;
									var attrValue = attrMatch.Groups[2].Value;
									attributes.Add($"{attrIndentStr}{attrName}=\"{attrValue}\"");
								}

								if(attributes.Count>0)
								{
									result.AppendLine($"{indentStr}<{tagName}");
									result.AppendLine(string.Join("\n",attributes));
									result.AppendLine($"{indentStr}{closing}");
									continue;
								}
							}
						}
					}

					// 普通行，保持原样
					result.AppendLine(line);
				}

				return result.ToString();
			} catch
			{
				// 如果格式化失败，返回原始 XML
				return xml;
			}
		}

		/// <summary>
		/// 带编码的 StringWriter
		/// </summary>
		private class StringWriterWithEncoding(Encoding encoding):StringWriter
		{
			private readonly Encoding _encoding = encoding;

			public override Encoding Encoding => _encoding;
		}

		/// <summary>
		/// 重新加载配置
		/// </summary>
		public static void Reload()
		{
			lock(_lock)
			{
				_instance=null;
			}
		}

		public void ApplyLoggingConfigToProfiler()
		{
			var logging = Logging;
			if(logging==null)
			{
				return;
			}

			PPA.Core.Profiler.EnableFileLogging=logging.EnableFileLogging;

			if(logging.MaxLogFiles>0)
			{
				PPA.Core.Profiler.MaxLogFiles=logging.MaxLogFiles;
			}

			if(logging.MaxLogAgeDays>0)
			{
				PPA.Core.Profiler.MaxLogAge=TimeSpan.FromDays(logging.MaxLogAgeDays);
			} else
			{
				PPA.Core.Profiler.MaxLogAge=null;
			}

			var minLevel = ParseMinimumLogLevel(logging.MinimumLogLevel);
			PPA.Core.Profiler.MinimumLogLevel=minLevel;
		}

		private static LogLevel ParseMinimumLogLevel(string value)
		{
			if(string.IsNullOrWhiteSpace(value))
			{
				return LogLevel.Information;
			}

			if(Enum.TryParse<LogLevel>(value,true,out var level))
			{
				return level;
			}

			return LogLevel.Information;
		}

		#endregion Singleton

		#region Configuration Properties

		/// <summary>
		/// 表格格式化配置
		/// </summary>
		[XmlElement("Table")]
		public TableFormattingConfig Table { get; set; } = new TableFormattingConfig();

		/// <summary>
		/// 文本格式化配置
		/// </summary>
		[XmlElement("Text")]
		public TextFormattingConfig Text { get; set; } = new TextFormattingConfig();

		/// <summary>
		/// 图表格式化配置
		/// </summary>
		[XmlElement("Chart")]
		public ChartFormattingConfig Chart { get; set; } = new ChartFormattingConfig();

		/// <summary>
		/// 快捷键配置
		/// </summary>
		[XmlElement("Shortcuts")]
		public ShortcutsConfig Shortcuts { get; set; } = new ShortcutsConfig();

		/// <summary>
		/// 日志配置
		/// </summary>
		[XmlElement("Logging")]
		public LoggingConfig Logging { get; set; } = new LoggingConfig();

		#endregion Configuration Properties
	}
}

using System.Drawing;
using NetOffice.OfficeApi.Enums;
using PPA.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// PPA 配置类 用于管理表格、文本、图表的格式化样式配置和快捷键配置
	/// </summary>
	[XmlRoot("PPAConfig")]
	public class FormattingConfig : PPA.Core.Abstraction.Business.IFormattingConfig
	{
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
						Profiler.LogMessage($"创建配置目录失败: {ex.Message}，使用用户目录");
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
				Profiler.LogMessage($"加载配置文件失败: {ex.Message}，使用默认配置");
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
				Profiler.LogMessage($"已加载配置文件: {filePath}");
				return config;
			} catch(Exception ex)
			{
				Profiler.LogMessage($"从文件加载配置失败: {ex.Message}");
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

				Profiler.LogMessage($"配置文件已保存: {configPath}");
			} catch(Exception ex)
			{
				Profiler.LogMessage($"保存配置文件失败: {ex.Message}");
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

		#endregion Singleton


		#region Table Formatting Configuration

		/// <summary>
		/// 表格格式化配置
		/// </summary>
		[XmlElement("Table")]
		public TableFormattingConfig Table { get; set; } = new TableFormattingConfig();

		#endregion Table Formatting Configuration

		#region Text Formatting Configuration

		/// <summary>
		/// 文本格式化配置
		/// </summary>
		[XmlElement("Text")]
		public TextFormattingConfig Text { get; set; } = new TextFormattingConfig();

		#endregion Text Formatting Configuration

		#region Chart Formatting Configuration

		/// <summary>
		/// 图表格式化配置
		/// </summary>
		[XmlElement("Chart")]
		public ChartFormattingConfig Chart { get; set; } = new ChartFormattingConfig();

		#endregion Chart Formatting Configuration

		#region Keyboard Shortcuts Configuration

		/// <summary>
		/// 快捷键配置
		/// </summary>
		[XmlElement("Shortcuts")]
		public ShortcutsConfig Shortcuts { get; set; } = new ShortcutsConfig();

		#endregion Keyboard Shortcuts Configuration
	}

	#region Table Formatting Config

	/// <summary>
	/// 表格格式化配置
	/// </summary>
	public class TableFormattingConfig
	{
		/// <summary>
		/// 表格样式 ID
		/// </summary>
		[XmlAttribute("StyleId")]
		public string StyleId { get; set; } = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";

		/// <summary>
		/// 数据行字体配置
		/// </summary>
		[XmlElement("DataRowFont")]
		public FontConfig DataRowFont { get; set; } = new FontConfig
		{
			Name="+mn-lt",
			NameFarEast="+mn-ea",
			Size=9.0f,
			Bold=false,
			ThemeColor="Dark1"
		};

		/// <summary>
		/// 标题行字体配置
		/// </summary>
		[XmlElement("HeaderRowFont")]
		public FontConfig HeaderRowFont { get; set; } = new FontConfig
		{
			Name="+mn-lt",
			NameFarEast="+mn-ea",
			Size=10.0f,
			Bold=true,
			ThemeColor="Dark1"
		};

		/// <summary>
		/// 数据行边框宽度（磅）
		/// </summary>
		[XmlAttribute("DataRowBorderWidth")]
		public float DataRowBorderWidth { get; set; } = 1.0f;

		/// <summary>
		/// 标题行边框宽度（磅）
		/// </summary>
		[XmlAttribute("HeaderRowBorderWidth")]
		public float HeaderRowBorderWidth { get; set; } = 1.75f;

		/// <summary>
		/// 数据行边框颜色主题
		/// </summary>
		[XmlAttribute("DataRowBorderColor")]
		public string DataRowBorderColor { get; set; } = "Accent2";

		/// <summary>
		/// 标题行边框颜色主题
		/// </summary>
		[XmlAttribute("HeaderRowBorderColor")]
		public string HeaderRowBorderColor { get; set; } = "Accent1";

		/// <summary>
		/// 是否启用数字格式化
		/// </summary>
		[XmlAttribute("AutoNumberFormat")]
		public bool AutoNumberFormat { get; set; } = true;

		/// <summary>
		/// 数字格式化保留的小数位数
		/// </summary>
		[XmlAttribute("DecimalPlaces")]
		public int DecimalPlaces { get; set; } = 0;

		/// <summary>
		/// 负数文本颜色（RGB 值）
		/// </summary>
		[XmlAttribute("NegativeTextColor")]
		public int NegativeTextColor { get; set; } = -65536; // 红色

		/// <summary>
		/// 表格全局设置
		/// </summary>
		[XmlElement("TableSettings")]
		public TableSettingsConfig TableSettings { get; set; } = new TableSettingsConfig();
	}

	/// <summary>
	/// 表格全局设置配置
	/// </summary>
	public class TableSettingsConfig
	{
		[XmlAttribute("FirstRow")]
		public bool FirstRow { get; set; } = true;

		[XmlAttribute("FirstCol")]
		public bool FirstCol { get; set; } = false;

		[XmlAttribute("LastRow")]
		public bool LastRow { get; set; } = false;

		[XmlAttribute("LastCol")]
		public bool LastCol { get; set; } = false;

		[XmlAttribute("HorizBanding")]
		public bool HorizBanding { get; set; } = false;

		[XmlAttribute("VertBanding")]
		public bool VertBanding { get; set; } = false;
	}

	#endregion Table Formatting Config

	#region Text Formatting Config

	/// <summary>
	/// 文本格式化配置
	/// </summary>
	public class TextFormattingConfig
	{
		/// <summary>
		/// 文本边距配置（厘米）
		/// </summary>
		[XmlElement("Margins")]
		public MarginsConfig Margins { get; set; } = new MarginsConfig
		{
			Top=0.2f,
			Bottom=0.2f,
			Left=0.5f,
			Right=0.5f
		};

		/// <summary>
		/// 字体配置
		/// </summary>
		[XmlElement("Font")]
		public FontConfig Font { get; set; } = new FontConfig
		{
			Name="+mn-lt",
			NameFarEast="+mn-ea",
			Size=16.0f,
			Bold=true,
			ThemeColor="Accent2"
		};

		/// <summary>
		/// 段落格式配置
		/// </summary>
		[XmlElement("Paragraph")]
		public ParagraphConfig Paragraph { get; set; } = new ParagraphConfig();

		/// <summary>
		/// 项目符号配置
		/// </summary>
		[XmlElement("Bullet")]
		public BulletConfig Bullet { get; set; } = new BulletConfig();

		/// <summary>
		/// 段落左缩进（厘米）
		/// </summary>
		[XmlAttribute("LeftIndent")]
		public float LeftIndent { get; set; } = 1.0f;
	}

	/// <summary>
	/// 边距配置
	/// </summary>
	public class MarginsConfig
	{
		[XmlAttribute("Top")]
		public float Top { get; set; }

		[XmlAttribute("Bottom")]
		public float Bottom { get; set; }

		[XmlAttribute("Left")]
		public float Left { get; set; }

		[XmlAttribute("Right")]
		public float Right { get; set; }
	}

	/// <summary>
	/// 字体配置
	/// </summary>
	public class FontConfig
	{
		[XmlAttribute("Name")]
		public string Name { get; set; } = "+mn-lt";

		[XmlAttribute("NameFarEast")]
		public string NameFarEast { get; set; } = "+mn-ea";

		[XmlAttribute("Size")]
		public float Size { get; set; }

		[XmlAttribute("Bold")]
		public bool Bold { get; set; }

		/// <summary>
		/// 主题颜色名称（如 "Dark1", "Accent1", "Accent2" 等）
		/// </summary>
		[XmlAttribute("ThemeColor")]
		public string ThemeColor { get; set; } = "Dark1";
	}

	/// <summary>
	/// 段落格式配置
	/// </summary>
	public class ParagraphConfig
	{
		[XmlAttribute("Alignment")]
		public string Alignment { get; set; } = "Justify";

		[XmlAttribute("WordWrap")]
		public bool WordWrap { get; set; } = true;

		[XmlAttribute("SpaceBefore")]
		public float SpaceBefore { get; set; } = 0;

		[XmlAttribute("SpaceAfter")]
		public float SpaceAfter { get; set; } = 0;

		[XmlAttribute("SpaceWithin")]
		public float SpaceWithin { get; set; } = 1.25f;

		[XmlAttribute("FarEastLineBreakControl")]
		public bool FarEastLineBreakControl { get; set; } = true;

		[XmlAttribute("HangingPunctuation")]
		public bool HangingPunctuation { get; set; } = true;
	}

	/// <summary>
	/// 项目符号配置
	/// </summary>
	public class BulletConfig
	{
		[XmlAttribute("Type")]
		public string Type { get; set; } = "Unnumbered";

		[XmlAttribute("Character")]
		public int Character { get; set; } = 9632; // 实心方块

		[XmlAttribute("FontName")]
		public string FontName { get; set; } = "Arial";

		[XmlAttribute("RelativeSize")]
		public float RelativeSize { get; set; } = 1.0f;

		[XmlAttribute("ThemeColor")]
		public string ThemeColor { get; set; } = "Dark1";
	}

	#endregion Text Formatting Config

	#region Chart Formatting Config

	/// <summary>
	/// 图表格式化配置
	/// </summary>
	public class ChartFormattingConfig
	{
		/// <summary>
		/// 常规字体配置
		/// </summary>
		[XmlElement("RegularFont")]
		public FontConfig RegularFont { get; set; } = new FontConfig
		{
			Name="+mn-lt",
			NameFarEast="+mn-ea",
			Size=8.0f,
			Bold=false,
			ThemeColor="Dark1"
		};

		/// <summary>
		/// 标题字体配置
		/// </summary>
		[XmlElement("TitleFont")]
		public FontConfig TitleFont { get; set; } = new FontConfig
		{
			Name="+mn-lt",
			NameFarEast="+mn-ea",
			Size=11.0f,
			Bold=true,
			ThemeColor="Dark1"
		};
	}

	#endregion Chart Formatting Config

	#region Keyboard Shortcuts Config

	/// <summary>
	/// 快捷键配置 格式：只配置数字或字母（如 "3", "C", "F1"），系统会自动添加 Ctrl 修饰键 空字符串表示不启用该快捷键 示例：FormatChart="3" 表示 Ctrl+3
	/// </summary>
	public class ShortcutsConfig
	{
		/// <summary>
		/// 美化表格快捷键（数字或字母，如 "1", "T"）
		/// </summary>
		[XmlAttribute("FormatTables")]
		public string FormatTables { get; set; } = string.Empty;

		/// <summary>
		/// 美化文本快捷键（数字或字母，如 "2", "X"）
		/// </summary>
		[XmlAttribute("FormatText")]
		public string FormatText { get; set; } = string.Empty;

		/// <summary>
		/// 美化图表快捷键（数字或字母，如 "3", "C"）
		/// </summary>
		[XmlAttribute("FormatChart")]
		public string FormatChart { get; set; } = string.Empty;

		/// <summary>
		/// 插入形状快捷键（数字或字母，如 "4", "I"）
		/// </summary>
		[XmlAttribute("CreateBoundingBox")]
		public string CreateBoundingBox { get; set; } = string.Empty;
	}

	#endregion Keyboard Shortcuts Config

	#region Helper Methods

	/// <summary>
	/// 配置辅助类 提供配置值的转换和辅助方法
	/// </summary>
	public static class ConfigHelper
	{
		/// <summary>
		/// 将主题颜色名称转换为 MsoThemeColorIndex
		/// </summary>
		public static MsoThemeColorIndex GetThemeColorIndex(string themeColorName)
		{
			if(string.IsNullOrEmpty(themeColorName))
				return MsoThemeColorIndex.msoThemeColorDark1;

			return themeColorName.ToLower() switch
			{
				"dark1" => MsoThemeColorIndex.msoThemeColorDark1,
				"dark2" => MsoThemeColorIndex.msoThemeColorDark2,
				"light1" => MsoThemeColorIndex.msoThemeColorLight1,
				"light2" => MsoThemeColorIndex.msoThemeColorLight2,
				"accent1" => MsoThemeColorIndex.msoThemeColorAccent1,
				"accent2" => MsoThemeColorIndex.msoThemeColorAccent2,
				"accent3" => MsoThemeColorIndex.msoThemeColorAccent3,
				"accent4" => MsoThemeColorIndex.msoThemeColorAccent4,
				"accent5" => MsoThemeColorIndex.msoThemeColorAccent5,
				"accent6" => MsoThemeColorIndex.msoThemeColorAccent6,
				"hyperlink" => MsoThemeColorIndex.msoThemeColorHyperlink,
				"followedhyperlink" => MsoThemeColorIndex.msoThemeColorFollowedHyperlink,
				_ => MsoThemeColorIndex.msoThemeColorDark1
			};
		}

		/// <summary>
		/// 将主题颜色名称转换为 RGB 颜色（0xRRGGBB）
		/// </summary>
		public static int GetThemeColorRgb(string themeColorName)
		{
			if(string.IsNullOrEmpty(themeColorName))
				return ColorTranslator.ToOle(Color.Black);

			Color color = themeColorName.ToLower() switch
			{
				"dark1" => Color.Black,
				"dark2" => ColorTranslator.FromHtml("#1F4E79"),
				"light1" => Color.White,
				"light2" => ColorTranslator.FromHtml("#F2F2F2"),
				"accent1" => ColorTranslator.FromHtml("#4472C4"),
				"accent2" => ColorTranslator.FromHtml("#ED7D31"),
				"accent3" => ColorTranslator.FromHtml("#A5A5A5"),
				"accent4" => ColorTranslator.FromHtml("#FFC000"),
				"accent5" => ColorTranslator.FromHtml("#5B9BD5"),
				"accent6" => ColorTranslator.FromHtml("#70AD47"),
				"hyperlink" => ColorTranslator.FromHtml("#0563C1"),
				"followedhyperlink" => ColorTranslator.FromHtml("#954F72"),
				_ => Color.Black
			};

			return ColorTranslator.ToOle(color);
		}

		/// <summary>
		/// 将段落对齐字符串转换为枚举
		/// </summary>
		public static NETOP.Enums.PpParagraphAlignment GetParagraphAlignment(string alignment)
		{
			if(string.IsNullOrEmpty(alignment))
				return NETOP.Enums.PpParagraphAlignment.ppAlignJustify;

			return alignment.ToLower() switch
			{
				"left" => NETOP.Enums.PpParagraphAlignment.ppAlignLeft,
				"center" => NETOP.Enums.PpParagraphAlignment.ppAlignCenter,
				"right" => NETOP.Enums.PpParagraphAlignment.ppAlignRight,
				"justify" => NETOP.Enums.PpParagraphAlignment.ppAlignJustify,
				"distribute" => NETOP.Enums.PpParagraphAlignment.ppAlignDistribute,
				_ => NETOP.Enums.PpParagraphAlignment.ppAlignJustify
			};
		}

		/// <summary>
		/// 将项目符号类型字符串转换为枚举
		/// </summary>
		public static NETOP.Enums.PpBulletType GetBulletType(string bulletType)
		{
			if(string.IsNullOrEmpty(bulletType))
				return NETOP.Enums.PpBulletType.ppBulletUnnumbered;

			return bulletType.ToLower() switch
			{
				"none" => NETOP.Enums.PpBulletType.ppBulletNone,
				"numbered" => NETOP.Enums.PpBulletType.ppBulletNumbered,
				"unnumbered" => NETOP.Enums.PpBulletType.ppBulletUnnumbered,
				"picture" => NETOP.Enums.PpBulletType.ppBulletPicture,
				_ => NETOP.Enums.PpBulletType.ppBulletUnnumbered
			};
		}

		/// <summary>
		/// 将厘米转换为磅（1 厘米 = 28.35 磅）
		/// </summary>
		public static float CmToPoints(float cm)
		{
			return cm*28.35f;
		}
	}

	#endregion Helper Methods
}

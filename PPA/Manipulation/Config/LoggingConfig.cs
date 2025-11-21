using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
	/// <summary>
	/// 日志配置
	/// </summary>
	public class LoggingConfig
	{
		/// <summary>
		/// 是否启用文件日志记录
		/// </summary>
		[XmlAttribute("EnableFileLogging")]
		public bool EnableFileLogging { get; set; } = true;

		/// <summary>
		/// 最多保留的日志文件数量
		/// </summary>
		[XmlAttribute("MaxLogFiles")]
		public int MaxLogFiles { get; set; } = 10;

		/// <summary>
		/// 日志文件最长保留天数
		/// </summary>
		[XmlAttribute("MaxLogAgeDays")]
		public int MaxLogAgeDays { get; set; } = 7;

		/// <summary>
		/// 最小写入日志级别（Debug、Information、Warning、Error）
		/// </summary>
		[XmlAttribute("MinimumLogLevel")]
		public string MinimumLogLevel { get; set; } = "Information";
	}
}

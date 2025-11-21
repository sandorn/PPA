using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
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
}

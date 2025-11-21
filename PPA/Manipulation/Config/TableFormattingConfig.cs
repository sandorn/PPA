using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
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
		/// 负数文本颜色（OLE RGB 值，255 表示红色）
		/// </summary>
		[XmlAttribute("NegativeTextColor")]
		public int NegativeTextColor { get; set; } = 255; // 红色 (BGR)

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
}

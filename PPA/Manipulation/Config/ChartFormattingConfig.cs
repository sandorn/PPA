using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
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
}

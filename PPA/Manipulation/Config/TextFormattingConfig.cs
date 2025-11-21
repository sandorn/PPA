using System.Xml.Serialization;

namespace PPA.Manipulation.Config
{
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
}

using NetOffice.OfficeApi.Enums;
using PPA.Core;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 文本格式化辅助类 提供文本形状的格式化功能
	/// </summary>
	internal static class TextFormatHelper
	{
		/// <summary>
		/// 应用文本格式化到指定形状
		/// </summary>
		/// <param name="shp"> 要格式化的形状对象 </param>
		internal static void ApplyTextFormatting(NETOP.Shape shp)
		{
			ExHandler.Run(() =>
			{
				// 从配置加载参数
				var config = FormattingConfig.Instance.Text;

				// 设置文本框边距（使用配置）
				var textFrame = shp.TextFrame;
				var margins = config.Margins;
				textFrame.MarginTop=ConfigHelper.CmToPoints(margins.Top);
				textFrame.MarginBottom=ConfigHelper.CmToPoints(margins.Bottom);
				textFrame.MarginLeft=ConfigHelper.CmToPoints(margins.Left);
				textFrame.MarginRight=ConfigHelper.CmToPoints(margins.Right);

				// 设置字体属性（使用配置）
				var tfont = textFrame.TextRange.Font;
				var fontConfig = config.Font;
				tfont.Name=fontConfig.Name;
				tfont.NameFarEast=fontConfig.NameFarEast;
				tfont.Color.ObjectThemeColor=ConfigHelper.GetThemeColorIndex(fontConfig.ThemeColor);
				tfont.Bold=fontConfig.Bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
				tfont.Size=fontConfig.Size;

				// 设置段落格式（使用配置）
				var paragraph = textFrame.TextRange.ParagraphFormat;
				var paraConfig = config.Paragraph;
				paragraph.FarEastLineBreakControl=paraConfig.FarEastLineBreakControl ? MsoTriState.msoTrue : MsoTriState.msoFalse;
				paragraph.HangingPunctuation=paraConfig.HangingPunctuation ? MsoTriState.msoTrue : MsoTriState.msoFalse;
				paragraph.BaseLineAlignment=NETOP.Enums.PpBaselineAlignment.ppBaselineAlignAuto;
				paragraph.Alignment=ConfigHelper.GetParagraphAlignment(paraConfig.Alignment);
				paragraph.WordWrap=paraConfig.WordWrap ? MsoTriState.msoTrue : MsoTriState.msoFalse;
				paragraph.SpaceBefore=paraConfig.SpaceBefore;
				paragraph.SpaceAfter=paraConfig.SpaceAfter;
				paragraph.SpaceWithin=paraConfig.SpaceWithin;

				// 设置项目符号（使用配置）
				var bullet = paragraph.Bullet;
				var bulletConfig = config.Bullet;
				bullet.Type=ConfigHelper.GetBulletType(bulletConfig.Type);
				bullet.Character=bulletConfig.Character;
				bullet.Font.Name=bulletConfig.FontName;
				bullet.RelativeSize=bulletConfig.RelativeSize;
				bullet.Font.Color.ObjectThemeColor=ConfigHelper.GetThemeColorIndex(bulletConfig.ThemeColor);

				// 设置悬挂缩进（使用配置）
				textFrame.Ruler.Levels[1].LeftMargin=ConfigHelper.CmToPoints(config.LeftIndent);
			});
		}
	}
}

using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;

namespace PPA.Formatting
{
	/// <summary>
	/// 文本格式化辅助类 提供文本形状的格式化功能
	/// </summary>
	/// <remarks>
	/// 构造函数，通过依赖注入获取配置
	/// </remarks>
	/// <param name="config">格式化配置</param>
	internal class TextFormatHelper(IFormattingConfig config) : ITextFormatHelper
	{
		private readonly IFormattingConfig _config = config ?? throw new System.ArgumentNullException(nameof(config));

		/// <summary>
		/// 应用文本格式化到指定形状
		/// </summary>
		/// <param name="shp"> 要格式化的形状对象 </param>
		public void ApplyTextFormatting(NETOP.Shape shp)
		{
			ExHandler.Run(() =>
			{
				// 从配置加载参数
				var config = _config.Text;

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

		/// <summary>
		/// 应用文本格式化到指定抽象形状
		/// </summary>
		/// <param name="shape"> 抽象形状对象 </param>
		public void ApplyTextFormatting(IShape shape)
		{
			PPA.Core.Profiler.LogMessage($"ApplyTextFormatting(IShape) 被调用，shape={shape?.GetType().Name ?? "null"}", "INFO");
			if(shape==null)
			{
				PPA.Core.Profiler.LogMessage("ApplyTextFormatting: shape 为 null，返回", "WARN");
				return;
			}

			// 检查 PowerPoint 适配器
			if(shape is PowerPointShape pptShape)
			{
				PPA.Core.Profiler.LogMessage("ApplyTextFormatting: 检测到 PowerPointShape，使用 PowerPoint 格式化", "INFO");
				ApplyTextFormatting(pptShape.NativeObject);
				return;
			}

			// 尝试从 IComWrapper 获取底层对象
			if(shape is IComWrapper<NETOP.Shape> typed)
			{
				PPA.Core.Profiler.LogMessage("ApplyTextFormatting: 检测到 IComWrapper<NETOP.Shape>，使用 PowerPoint 格式化", "INFO");
				ApplyTextFormatting(typed.NativeObject);
				return;
			}

			var native = (shape as IComWrapper)?.NativeObject as NETOP.Shape;
			if(native!=null)
			{
				PPA.Core.Profiler.LogMessage("ApplyTextFormatting: 检测到 NetOffice.Shape，使用 PowerPoint 格式化", "INFO");
				ApplyTextFormatting(native);
				return;
			}

			PPA.Core.Profiler.LogMessage($"ApplyTextFormatting: 未知的形状类型 {shape.GetType().FullName}，无法格式化", "ERROR");
		}


		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

		private static void SafeSet(System.Action action)
		{
			try { action(); } catch { }
		}
	}
}

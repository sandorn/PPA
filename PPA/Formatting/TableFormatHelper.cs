using System;
using NetOffice.OfficeApi.Enums;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;

namespace PPA.Formatting
{
	/// <summary>
	/// 表格格式化辅助类 提供表格的高性能格式化功能
	/// </summary>
	/// <remarks>
	/// 构造函数，通过依赖注入获取配置
	/// </remarks>
	/// <param name="config">格式化配置</param>
	internal class TableFormatHelper(IFormattingConfig config): ITableFormatHelper
	{
		private static readonly int NegativeTextColor = ColorTranslator.ToOle(Color.Red);
		private readonly IFormattingConfig _config = config ?? throw new System.ArgumentNullException(nameof(config));

		/// <summary>
		/// 对表格进行高性能格式化。
		/// </summary>
		/// <param name="tbl"> 要格式化的 PowerPoint 表格对象。 </param>
		public void FormatTables(NETOP.Table tbl)
		{
			// 从配置加载参数
			var tableConfig = _config.Table;

			string styleId = tableConfig.StyleId;
			MsoThemeColorIndex dk1 = ConfigHelper.GetThemeColorIndex(tableConfig.DataRowFont.ThemeColor);
			MsoThemeColorIndex a1 = ConfigHelper.GetThemeColorIndex(tableConfig.HeaderRowBorderColor);
			MsoThemeColorIndex a2 = ConfigHelper.GetThemeColorIndex(tableConfig.DataRowBorderColor);

			float thin = tableConfig.DataRowBorderWidth;
			float thick = tableConfig.HeaderRowBorderWidth;
			float fontSize = tableConfig.DataRowFont.Size;
			float bigFontSize = tableConfig.HeaderRowFont.Size;
			string fontName = tableConfig.DataRowFont.Name;
			string fontNameFarEast = tableConfig.DataRowFont.NameFarEast;

			bool useAutoNum = tableConfig.AutoNumberFormat;
			int decimalPlacesValue = tableConfig.DecimalPlaces;

			int rows = tbl.Rows.Count;
			int cols = tbl.Columns.Count;

			// --- 2. 一次性设置表格全局样式（使用配置） ---
			tbl.ApplyStyle(styleId,false);
			var tableSettings = tableConfig.TableSettings;
			tbl.FirstRow=tableSettings.FirstRow;
			tbl.FirstCol=tableSettings.FirstCol;
			tbl.LastRow=tableSettings.LastRow;
			tbl.LastCol=tableSettings.LastCol;
			tbl.HorizBanding=tableSettings.HorizBanding;
			tbl.VertBanding=tableSettings.VertBanding;

			// --- 3. 性能优化：批处理模式 --- 预先创建批处理集合
			var firstRowCells = new List<NETOP.Cell>();
			var lastRowCells = new List<NETOP.Cell>();
			var dataRowCells = new List<NETOP.Cell>();

			// 第一步：收集所有单元格到不同集合
			for(int r = 1;r<=rows;r++)
			{
				var row = tbl.Rows[r];
				for(int c = 1;c<=cols;c++)
				{
					var cell = row.Cells[c];
					dataRowCells.Add(cell);
					// 只收集引用，不立即处理
					if(r==1)
						firstRowCells.Add(cell);
					else if(r==rows)
						lastRowCells.Add(cell);
				}
			}

			//批量处理数据行
			FormatDataRowCells(dataRowCells,fontName,fontNameFarEast,fontSize,dk1,thin,a2,useAutoNum,decimalPlacesValue,tableConfig.NegativeTextColor);

			//批量处理标题行和尾行
			FormatOutsideRowCells(firstRowCells,lastRowCells,
				tableConfig.HeaderRowFont.Name,
				tableConfig.HeaderRowFont.NameFarEast,
				bigFontSize,
				dk1,
				thick,
				a1);
		}

		/// <summary>
		/// 对表格进行高性能格式化（抽象接口版本）。
		/// </summary>
		/// <param name="tbl"> 要格式化的抽象表格对象。 </param>
		public void FormatTables(ITable tbl)
		{
			PPA.Core.Profiler.LogMessage($"FormatTables(ITable) 被调用，tbl={tbl?.GetType().Name ?? "null"}", "INFO");
			if(tbl==null)
			{
				PPA.Core.Profiler.LogMessage("FormatTables: tbl 为 null，返回", "WARN");
				return;
			}

			// 检查 PowerPoint 适配器
			if(tbl is PowerPointTable pptTable)
			{
				PPA.Core.Profiler.LogMessage("FormatTables: 检测到 PowerPointTable，使用 PowerPoint 格式化", "INFO");
				// PowerPointTable 包装了 NETOP.Table，需要获取它
				var native = (pptTable as IComWrapper)?.NativeObject as NETOP.Table;
				if(native != null)
				{
					FormatTables(native);
					return;
				}
			}

			// 尝试从抽象表格中获取底层 NetOffice 对象
			var nativeTable = (tbl as IComWrapper)?.NativeObject as NETOP.Table;
			if(nativeTable!=null)
			{
				PPA.Core.Profiler.LogMessage("FormatTables: 检测到 NetOffice.Table，使用 PowerPoint 格式化", "INFO");
				FormatTables(nativeTable);
				return;
			}

			// 兜底：尝试已知的强类型包装
			if(tbl is IComWrapper<NETOP.Table> typed)
			{
				PPA.Core.Profiler.LogMessage("FormatTables: 检测到 IComWrapper<NETOP.Table>，使用 PowerPoint 格式化", "INFO");
				FormatTables(typed.NativeObject);
				return;
			}

			PPA.Core.Profiler.LogMessage($"FormatTables: 未知的表格类型 {tbl.GetType().FullName}，无法格式化", "ERROR");
		}


		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

		private static int AdjustColorBrightness(int oleColor, double factor)
		{
			factor = Math.Max(-1.0, Math.Min(1.0, factor));

			Color baseColor = ColorTranslator.FromOle(oleColor);
			double r = baseColor.R;
			double g = baseColor.G;
			double b = baseColor.B;

			if(factor >= 0)
			{
				r = r + ((255 - r) * factor);
				g = g + ((255 - g) * factor);
				b = b + ((255 - b) * factor);
			}
			else
			{
				double scale = 1 + factor;
				r *= scale;
				g *= scale;
				b *= scale;
			}

			var adjusted = Color.FromArgb(
				(int)Math.Round(Math.Max(0, Math.Min(255, r))),
				(int)Math.Round(Math.Max(0, Math.Min(255, g))),
				(int)Math.Round(Math.Max(0, Math.Min(255, b))));

			return ColorTranslator.ToOle(adjusted);
		}

		private static bool TryFormatNumericValue(string text, TableFormattingConfig config, out string formatted, out bool isNegative)
		{
			formatted = text;
			isNegative = false;

			if(string.IsNullOrWhiteSpace(text))
				return false;

			string trimmed = text.Trim();

			if(double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double value) ||
				double.TryParse(trimmed, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value))
			{
				isNegative = value < 0;
				string formatSpecifier = config.DecimalPlaces > 0 ? $"N{config.DecimalPlaces}" : "N0";
				formatted = value.ToString(formatSpecifier, CultureInfo.CurrentCulture);
				return true;
			}

			return false;
		}

		private static void SafeSet(System.Action action)
		{
			try { action(); } catch { }
		}

		/// <summary>
		/// 批量处理首末行的单元格，减少重复操作和COM调用
		/// </summary>
		private void FormatOutsideRowCells(List<NETOP.Cell> firstRowCells,List<NETOP.Cell> lastRowCells,string fontName,string fontNameFarEast,float fontSize,MsoThemeColorIndex txtColor,float borderWidth,MsoThemeColorIndex borderColor)
		{
			// 设置首行上下边框
			for(int i = 0;i<firstRowCells.Count;i++)
			{
				var cell = firstRowCells[i];
				cell.Shape.Fill.Visible=MsoTriState.msoFalse;
				var textRange = cell.Shape.TextFrame.TextRange;
				SetFontProperties(textRange,fontName,fontNameFarEast,fontSize,MsoTriState.msoTrue,txtColor);

				textRange.ParagraphFormat.Alignment=NETOP.Enums.PpParagraphAlignment.ppAlignCenter;

				// 边框
				SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,borderWidth,(object) borderColor);
				SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,borderWidth,(object) borderColor);
			}

			// 设置末行下边框
			for(int i = 0;i<lastRowCells.Count;i++)
			{
				var cell = lastRowCells[i];
				SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,borderWidth,(object) borderColor);
			}
		}

		/// <summary>
		/// 批量处理数据行的单元格，使用更高效的处理方式
		/// </summary>
		private void FormatDataRowCells(List<NETOP.Cell> cells,string fontName,string fontNameFarEast,float fontSize,MsoThemeColorIndex txtColor,float thinBorderWidth,MsoThemeColorIndex thinBorderColor,bool autonum,int decimalPlaces,int negativeTextColor)
		{
			int cellCount = cells.Count;

			for(int i = 0;i<cellCount;i++)
			{
				var cell = cells[i];
				cell.Shape.Fill.Visible=MsoTriState.msoFalse;

				var textRange = cell.Shape.TextFrame.TextRange;
				SetFontProperties(textRange,fontName,fontNameFarEast,fontSize,MsoTriState.msoFalse,txtColor);

				// 智能优化：只对非空文本进行数字格式化
				if(autonum&&!string.IsNullOrEmpty(textRange.Text.Trim()))
				{
					SmartNumberFormat(textRange,decimalPlaces,negativeTextColor);
				}

				SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,thinBorderWidth,(object) thinBorderColor,0.5f);
			}
		}

		/// <summary>
		/// 批量设置字体属性，减少 COM 调用次数。
		/// </summary>
		private void SetFontProperties(NETOP.TextRange textRange,string name,string nameFarEast,float size,MsoTriState bold,MsoThemeColorIndex color)
		{
			// 关键：通过 .Font 来访问字体属性
			textRange.Font.Name=name;
			textRange.Font.NameFarEast=nameFarEast;
			textRange.Font.Size=size;
			textRange.Font.Bold=bold;
			textRange.Font.Color.ObjectThemeColor=color;
		}

		/// <summary>
		/// 高性能数字格式化，针对大量单元格优化，在必要时修改文本和颜色
		/// </summary>
		private void SmartNumberFormat(NETOP.TextRange textRange,int decimalPlaces,int negativeTextColor)
		{
			// 性能优化1: 直接访问文本，避免多次Trim操作
			string text = textRange.Text;
			if(string.IsNullOrEmpty(text)) return;

			// 预先计算可能的百分比符号位置
			int length = text.Length;
			bool isPercentage = length > 0 && text[length - 1] == '%';

			// 获取需要解析的数字部分
			string numStr = isPercentage ? text.Substring(0, length - 1).Trim() : text.Trim();

			// 性能优化2: 快速检查是否可能是数字
			if(string.IsNullOrEmpty(numStr)||
				(!char.IsDigit(numStr[0])&&numStr[0]!='-'&&numStr[0]!='.'&&numStr[0]!='+'))
			{
				return;
			}

			// 性能优化3: 尝试解析数字
			if(!double.TryParse(numStr,NumberStyles.Any,CultureInfo.InvariantCulture,out double num))
			{
				return;
			}

			// 性能优化4: 预缓存常用格式字符串
			string format = decimalPlaces switch
			{
				0 => "N0",
				1 => "N1",
				2 => "N2",
				3 => "N3",
				_ => "N" + decimalPlaces,
			};
			string formatted = num.ToString(format);
			if(isPercentage)
			{
				formatted+="%";
			}

			// 性能优化5: 避免不必要的COM调用 - 只有当文本真的需要改变时才设置
			if(text!=formatted)
			{
				textRange.Text=formatted;
			}

			// 性能优化6: 负数颜色设置 - 只在需要时调用
			if(num<0)
			{
				textRange.Font.Color.RGB=negativeTextColor;
			}
		}

		/// <summary>
		/// 设置单元格边框
		/// </summary>
		private void SetBorder(NETOP.Cell cell,NETOP.Enums.PpBorderType borderType,float setWeight,object tcolor,float transparency = 0)
		{
			var border = cell.Borders[borderType];

			// weight 为 0,隐藏条线
			if(setWeight<=0f)
			{
				border.Weight=setWeight;
				border.Visible=MsoTriState.msoFalse;
			} else
			{
				border.Weight=setWeight;
				border.Visible=MsoTriState.msoTrue;
				
				// WPS 不支持 Transparency，需要安全设置
				try
				{
					border.Transparency=transparency;
				}
				catch(System.Exception ex)
				{
					// WPS 可能不支持 Transparency，忽略错误
					PPA.Core.Profiler.LogMessage($"设置边框透明度失败（可能是不支持的属性）: {ex.Message}", "WARN");
				}
				
				// 使用模式匹配简化颜色逻辑
				if(tcolor is MsoThemeColorIndex themeColor) border.ForeColor.ObjectThemeColor=themeColor;
				else if(tcolor is int rgbColor) border.ForeColor.RGB=rgbColor;
				else border.ForeColor.RGB=0; // 默认黑色
			}
		}
	}
}

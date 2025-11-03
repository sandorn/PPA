using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using ToastAPI;
using VBAApi;

using NETOP = NetOffice.PowerPointApi;

namespace PPA.Helpers
{
	public static class FormatHelper
	{
		private static readonly int NegativeTextColor = ColorTranslator.ToOle(Color.Red);

		#region Internal Methods

		internal static void ApplyTextFormatting(NETOP.Shape shp)
		{
			ExHandler.Run(() =>
			{
				const MsoThemeColorIndex tcolor = MsoThemeColorIndex.msoThemeColorText1;
				const MsoThemeColorIndex acolor = MsoThemeColorIndex.msoThemeColorAccent2;

				// 设置文本框边距
				var textFrame = shp.TextFrame;
				textFrame.MarginTop = 0.2f * 28.35f;   // 上边距 0.2cm → 磅
				textFrame.MarginBottom = 0.2f * 28.35f; // 下边距
				textFrame.MarginLeft = 0.5f * 28.35f;   // 左边距 0.5cm → 磅
				textFrame.MarginRight = 0.5f * 28.35f;  // 右边距

				// 设置字体属性
				var tfont = textFrame.TextRange.Font;
				tfont.Name = "+mn-lt";
				tfont.NameFarEast = "+mn-ea";
				tfont.Color.ObjectThemeColor = acolor;
				tfont.Bold = MsoTriState.msoTrue;
				tfont.Size = 16f;

				// 设置段落格式
				var paragraph = textFrame.TextRange.ParagraphFormat;
				paragraph.FarEastLineBreakControl = MsoTriState.msoTrue;
				paragraph.HangingPunctuation = MsoTriState.msoTrue;
				paragraph.BaseLineAlignment = NETOP.Enums.PpBaselineAlignment.ppBaselineAlignAuto;
				paragraph.Alignment = NETOP.Enums.PpParagraphAlignment.ppAlignJustify;
				paragraph.WordWrap = MsoTriState.msoTrue;
				paragraph.SpaceBefore = 0;
				paragraph.SpaceAfter = 0;
				paragraph.SpaceWithin = 1.25f;

				// 设置项目符号
				var bullet = paragraph.Bullet;
				bullet.Type = NETOP.Enums.PpBulletType.ppBulletUnnumbered;
				//bullet.Character = 9632; // 实心方块
				bullet.Font.Name = "Arial";
				bullet.RelativeSize = 1.0f;
				bullet.Font.Color.ObjectThemeColor = tcolor;

				// 设置悬挂缩进（通过 Ruler 对象）
				textFrame.Ruler.Levels[1].LeftMargin = 1.0f * 28.35f; // 厘米转磅,段落左缩进
			},"格式化文本框字体");
		}

		internal static void FormatChartText(NETOP.Shape shape)
		{
			var chart = shape.Chart;

			// 设置字体
			const float size = 8f;
			const float titleSize = 11f;

			// 设置图表标题字体
			if(chart.HasTitle)
			{
				chart.ChartTitle.Font.Name = "+mn-lt";
				chart.ChartTitle.Font.Bold = MsoTriState.msoTrue;
				chart.ChartTitle.Font.Size = titleSize;
			}

			// 设置图例字体
			if(chart.HasLegend)
			{
				chart.Legend.Font.Name = "+mn-lt";
				chart.Legend.Font.Size = size;
			}

			// 设置数据表字体
			if(chart.HasDataTable)
			{
				chart.DataTable.Font.Name = "+mn-lt";
				chart.DataTable.Font.Size = size;
			}

			// 设置数据标签字体
			dynamic seriesCollection = chart.SeriesCollection();
			foreach(var series in seriesCollection)
			{
				ExHandler.Run(() =>
				{
					var points = series.Points();
					foreach(var point in points)
					{
						if(point.HasDataLabel)
						{
							point.DataLabel.Font.Name = "+mn-lt";
							point.DataLabel.Font.Size = size;
						}
					}
				},"格式化图表字体");
			}

			// 设置坐标轴字体
			ExHandler.Run(() =>
			{
				// 获取图表类型并检查是否为不支持坐标轴的类型
				var chartType = chart.ChartType;
				//不支持坐标轴的图表类型
				var nonAxisCharts = new HashSet<XlChartType>{
						XlChartType.xlPie,        // 饼图
						XlChartType.xl3DPie,      // 3D饼图
						XlChartType.xlDoughnut,   // 圆环图
						XlChartType.xlPieOfPie,   // 复合饼图
						XlChartType.xlBarOfPie,   // 条形饼图
						XlChartType.xlRadar,      // 雷达图
						XlChartType.xlRadarFilled // 填充雷达图
					};
				if(nonAxisCharts.Contains(chartType))
				{
					Debug.WriteLine($"图表类型 {chartType} 不支持坐标轴，已跳过");
					return;
				}

				// 安全设置各坐标轴
				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlSecondary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlSecondary,size);
			},"格式化图表字体");
		}

		/// <summary>
		/// 对表格进行极限优化的格式化。
		/// </summary>
		/// <param name="tbl">要格式化的 PowerPoint 表格对象。</param>
		/// <param name="autonum">是否自动格式化数字。</param>
		/// <param name="decimalPlaces">保留的小数位数。</param>
		internal static void FormatTables(NETOP.Table tbl,bool autonum = true,int decimalPlaces = 0)
		{
			//const string styleId = "{5940675A-B579-460E-94D1-54222C63F5DA}"; //styleName="无样式，网格型">
			//const string styleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"; //styleName="中度样式 2 - 强调 1">
			//const string styleId = "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}"; //styleName="浅色样式 2 - 强调 2">
			//const string styleId = "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}"; // styleName="浅色样式 2 - 强调 1">
			// --- 1. 预定义所有常量 ---
			const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
			const MsoThemeColorIndex txtColor = MsoThemeColorIndex.msoThemeColorText1;
			const MsoThemeColorIndex bdColor1 = MsoThemeColorIndex.msoThemeColorAccent1;
			const MsoThemeColorIndex bdColor2 = MsoThemeColorIndex.msoThemeColorAccent2;

			const float thin = 1.0f, thick = 2.0f;
			const float fontSize = 9.0f, bigFontSize = 10.0f;
			const string fontName = "+mn-lt";
			const string fontNameFarEast = "+mn-ea";

			int rows = tbl.Rows.Count;
			int cols = tbl.Columns.Count;

			// --- 2. 一次性设置表格全局样式 ---
			tbl.ApplyStyle(styleId,false);
			tbl.FirstRow=true;
			tbl.FirstCol=false;
			tbl.LastRow=false;
			tbl.LastCol=false;
			tbl.HorizBanding=false;
			tbl.VertBanding=false;

			// --- 3. 核心循环：使用 NetOffice 的 Dispose 模式 ---
			for(int r = 1;r<=rows;r++)
			{
				// 使用 using 语句，这是处理 IDisposable 对象的最佳实践
				using(var row = tbl.Rows[r])
				{
					bool isFirstRow = (r == 1);
					bool isLastRow = (r == rows);

					float topBorderWidth = isFirstRow ? thick : thin;
					var topBorderColor_obj = (object)(isFirstRow ? bdColor1 : bdColor2);
					float bottomBorderWidth = isLastRow ? thick : thin;
					var bottomBorderColor_obj = (object)bdColor1;

					for(int c = 1;c<=cols;c++)
					{
						// 嵌套 using 语句，确保每个对象都被及时释放
						using(var cell = row.Cells[c])
						{
							cell.Shape.Fill.Visible=MsoTriState.msoFalse;
							using(var textRange = cell.Shape.TextFrame.TextRange)
							{
								if(isFirstRow)
								{
									SetFontProperties(textRange,fontName,fontNameFarEast,bigFontSize,MsoTriState.msoTrue,txtColor);
									textRange.ParagraphFormat.Alignment=NETOP.Enums.PpParagraphAlignment.ppAlignCenter;
								} else
								{
									SetFontProperties(textRange,fontName,fontNameFarEast,fontSize,MsoTriState.msoFalse,txtColor);
									if(autonum)
									{
										SmartNumberFormat(textRange,decimalPlaces,NegativeTextColor);
									}
								}
							} // textRange.Dispose() 在这里被自动调用

							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,topBorderWidth,topBorderColor_obj);
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,bottomBorderWidth,bottomBorderColor_obj);
						} // cell.Dispose() 在这里被自动调用
					}
				} // row.Dispose() 在这里被自动调用
			}
		}


		/// <summary>
		/// 批量设置字体属性，减少 COM 调用次数。
		/// </summary>
		private static void SetFontProperties(NETOP.TextRange textRange,string name,string nameFarEast,float size,MsoTriState bold,MsoThemeColorIndex color)
		{
			// 关键：通过 .Font 来访问字体属性
			textRange.Font.Name=name;
			textRange.Font.NameFarEast=nameFarEast;
			textRange.Font.Size=size;
			textRange.Font.Bold=bold;
			textRange.Font.Color.ObjectThemeColor=color;
		}

		/// <summary>
		/// 智能格式化数字，只在必要时修改文本和颜色。
		/// </summary>
		private static void SmartNumberFormat(NETOP.TextRange textRange,int decimalPlaces,int negativeTextColor)
		{
			var original = textRange.Text.Trim();
			if(string.IsNullOrEmpty(original)) return;

			var isPercentage = original.EndsWith("%");
			var numStr = isPercentage ? original.Substring(0, original.Length - 1) : original;

			// 使用 InvariantCulture 来确保小数点解析的正确性
			if(!double.TryParse(numStr,NumberStyles.Any,CultureInfo.InvariantCulture,out var num))
			{
				return;
			}

			// 构造格式字符串，例如 "N0", "N2"
			var format = $"N{decimalPlaces}";
			var formatted = num.ToString(format);

			if(isPercentage)
				formatted+="%";

			// 关键优化：只有当文本真的需要改变时才设置，避免触发不必要的重绘和缓存刷新
			if(textRange.Text!=formatted)
			{
				textRange.Text=formatted;
			}

			// 负数颜色设置
			if(num<0)
			{
				textRange.Font.Color.RGB=negativeTextColor;
			}
		}

		internal static void FormatTablesbyVBA(NETOP.Application app,NETOP.Slide slide)
		{
			if(app==null||slide==null)
			{
				Toast.Show("无效的应用程序或幻灯片对象",Toast.ToastType.Error);
				return;
			}

			const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
			int tableCount = 0;

			// 先应用表格样式，并统计有多少个表格
			foreach(NETOP.Shape shape in slide.Shapes)
			{
				if(shape.HasTable==MsoTriState.msoTrue)
				{
					NETOP.Table tbl = shape.Table;
					tbl.FirstRow=true;    // 标题行（首行特殊格式）
					tbl.FirstCol=false;   // 标题列（首列特殊格式）
					tbl.LastRow=false;    // 汇总行（末行特殊格式）
					tbl.LastCol=false;    // 汇总列（末列特殊格式）
					tbl.HorizBanding=false;   //镶边行
					tbl.VertBanding=false;    //镶边列
					tbl.ApplyStyle(styleId,false);
					tableCount++;
				}
			}

			if(tableCount==0)
			{
				Toast.Show("当前幻灯片上没有表格",Toast.ToastType.Info);
				return;
			}

			// 显示等待光标
			System.Windows.Forms.Cursor.Current=System.Windows.Forms.Cursors.WaitCursor;

			// 调用新的管理器，它会自动处理模块初始化
			VbaManager.RunMacro(app,"FormatAllTables");
			Toast.Show($"成功格式化了 {tableCount} 个表格",Toast.ToastType.Success);

			// 恢复光标
			System.Windows.Forms.Cursor.Current=System.Windows.Forms.Cursors.Default;
		}

		#endregion Internal Methods

		#region Private Methods

		private static void SafeSetAxis(NETOP.Chart chart,XlAxisType axisType,XlAxisGroup axisGroup,float size)
		{
			NETOP.Axis axis = null;
			ExHandler.Run(() =>
			{
				axis = (NETOP.Axis) chart.Axes(axisType,axisGroup);
				if(ShapeUtils.IsInvalidComObject(axis)) axis = null;
			},"获取坐标轴");

			if(axis == null) return;

			// 刻度标签
			ExHandler.Run(() =>
			{
				if(axis.TickLabels != null)
				{
					axis.TickLabels.Font.Name = "+mn-lt";
					axis.TickLabels.Font.Size = size;
				}
			},"设置坐标轴刻度标签");

			// 坐标轴标题
			ExHandler.Run(() =>
			{
				var hasTitle = axis.HasTitle;
				if(!hasTitle || axis.AxisTitle == null) return;
				axis.AxisTitle.Font.Name = "+mn-lt";
				axis.AxisTitle.Font.Size = size;
			},"设置坐标轴标题");
		}

		// 判断 tcolor 类型
		private static void SetBorder(NETOP.Cell cell,NETOP.Enums.PpBorderType borderType,float setWeight,object tcolor)
		{
			var border = cell.Borders[borderType];

			// weight 为 0,隐藏条线
			if(setWeight <= 0f)
			{
				border.Weight = setWeight;
				border.Visible = MsoTriState.msoFalse;
			} else
			{
				border.Weight = setWeight;
				border.Visible = MsoTriState.msoTrue;
				border.Transparency = 0f;
				// 使用模式匹配简化颜色逻辑
				if(tcolor is MsoThemeColorIndex themeColor) border.ForeColor.ObjectThemeColor = themeColor;
				else if(tcolor is int rgbColor) border.ForeColor.RGB = rgbColor;
				else border.ForeColor.RGB = 0; // 默认黑色
			}
		}
		
		#endregion Private Methods
	}
}
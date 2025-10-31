using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using VBAApi;

using NETOP = NetOffice.PowerPointApi;

namespace PPA.Helpers
{
	public static class FormatHelper
	{
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

		internal static void FormatTables(NETOP.Table tbl,bool autonum = true,int decimalPlaces = 0)
		{
			const MsoThemeColorIndex txtcolor = MsoThemeColorIndex.msoThemeColorText1;
			const MsoThemeColorIndex bdcolor1 = MsoThemeColorIndex.msoThemeColorAccent1;
			const MsoThemeColorIndex bdcolor2 = MsoThemeColorIndex.msoThemeColorAccent2;
			//const string styleId = "{5940675A-B579-460E-94D1-54222C63F5DA}"; //styleName="无样式，网格型">
			//const string styleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"; //styleName="中度样式 2 - 强调 1">
			//const string styleId = "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}"; //styleName="浅色样式 2 - 强调 2">
			//const string styleId = "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}"; // styleName="浅色样式 2 - 强调 1">
			const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"; //  styleName = "浅色样式 1 - 强调 1" >
			tbl.FirstRow = true;    // 标题行（首行特殊格式）
			tbl.FirstCol = false;   // 标题列（首列特殊格式）
			tbl.LastRow = false;    // 汇总行（末行特殊格式）
			tbl.LastCol = false;    // 汇总列（末列特殊格式）
			tbl.HorizBanding = false;    //镶边行
			tbl.VertBanding = false;    //镶边列
			tbl.ApplyStyle(styleId,false);

			const float thin = 1.0f, thick = 2.0f;
			const float fontsize = 9.0f, bigfontsize = 10.0f;
			int rows = tbl.Rows.Count, cols = tbl.Columns.Count;

			ExHandler.Run(() =>
			{
				for(var r = 1;r <= rows;r++)
				{
					bool isFirstR = (r == 1);
					bool isLastR = (r == rows);

					var row = tbl.Rows[r];
					foreach(NETOP.Cell cell in row.Cells.Cast<NETOP.Cell>())
					{
						//cell.Shape.Fill.ForeColor.ObjectThemeColor = frbgcolor; //填充颜色
						// 清除填充
						cell.Shape.Fill.Visible = MsoTriState.msoFalse;

						var rng = cell.Shape.TextFrame.TextRange;
						rng.Font.Name = "+mn-lt";
						rng.Font.NameFarEast = "+mn-ea"; // 设置为中文“正文字体” ;

						rng.Font.Size = isFirstR ? bigfontsize : fontsize;
						rng.Font.Bold = isFirstR ? MsoTriState.msoTrue : MsoTriState.msoFalse;
						rng.Font.Color.ObjectThemeColor = txtcolor;
						if(isFirstR) rng.ParagraphFormat.Alignment = NETOP.Enums.PpParagraphAlignment.ppAlignCenter;
						if(!isFirstR && autonum) SmartNumberFormat(rng,decimalPlaces); // 数值智能格式

						if(isFirstR)
						{
							// 首行顶线底线
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,thick,bdcolor1);
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,thick,bdcolor1);
						} else if(isLastR)
						{
							// 尾行底线
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,thin,bdcolor2);
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,thick,bdcolor1);
						} else
						{
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderTop,thin,bdcolor2);
							SetBorder(cell,NETOP.Enums.PpBorderType.ppBorderBottom,thin,bdcolor2);
						}
					}
				}
			},"格式化表格");
		}

		internal static void FormatTablesbyVBA(NETOP.Application app,NETOP.Slide slide)
		{
			const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
			foreach(NETOP.Shape shape in slide.Shapes)
			{
				if(shape.HasTable == MsoTriState.msoTrue)
				{
					NETOP.Table tbl = shape.Table;
					tbl.FirstRow = true;    // 标题行（首行特殊格式）
					tbl.FirstCol = false;   // 标题列（首列特殊格式）
					tbl.LastRow = false;    // 汇总行（末行特殊格式）
					tbl.LastCol = false;    // 汇总列（末列特殊格式）
					tbl.HorizBanding = false;   //镶边行
					tbl.VertBanding = false;    //镶边列
					tbl.ApplyStyle(styleId,false);
				}
			}

			const string vbaCode =
				@"Sub FormatAllTables()
					On Error Resume Next
					Dim sld As Slide
					Dim shp As Shape

					' 获取当前页面
					Set sld = ActiveWindow.View.Slide

					' 遍历页面中的所有形状
					For Each shp In sld.Shapes
						If shp.HasTable Then
							FormatSingleTable shp.Table
						End If
					Next shp
				End Sub

				Private Sub FormatSingleTable(tbl As Table)
					Const txtColor As Long = msoThemeColorText1
					Const bdColor1 As Long = msoThemeColorAccent1
					Const bdColor2 As Long = msoThemeColorAccent2

					Const thin As Single = 1#
					Const thick As Single = 2#
					Const fontSize As Single = 9#
					Const bigFontSize As Single = 10#

					Dim rows As Long: rows = tbl.Rows.Count
					Dim cols As Long: cols = tbl.Columns.Count
					Dim r As Long, c As Long
					Dim cell As Cell
					Dim txtRng As TextRange

					' ===== 主要优化点 =====
					' 首行特殊处理（减少循环内判断）
					For c = 1 To cols
						Set cell = tbl.Cell(1, c)
						' 清除填充色
						cell.Shape.Fill.Visible = msoFalse
						Set txtRng = cell.Shape.TextFrame.TextRange

						With txtRng
							.Font.Name = ""+mn-lt""
							.Font.NameFarEast = ""+mn-ea""
							.Font.Size = bigFontSize
							.Font.Bold = msoTrue
							.Font.Color.ObjectThemeColor = txtColor
							.ParagraphFormat.Alignment = ppAlignCenter
						End With

						' 设置首行特殊边框
						With cell.Borders(ppBorderTop)
							.Weight = thick
							.ForeColor.ObjectThemeColor = bdColor1
						End With
						With cell.Borders(ppBorderBottom)
							.Weight = thick
							.ForeColor.ObjectThemeColor = bdColor1
						End With
					Next c

					' 其他行处理（优化循环）
					For r = 2 To rows
						For c = 1 To cols
							Set cell = tbl.Cell(r, c)
							cell.Shape.Fill.Visible = msoFalse
							Set txtRng = cell.Shape.TextFrame.TextRange

							' 设置通用格式
							With txtRng
								.Font.Name = ""+mn-lt""
								.Font.NameFarEast = ""+mn-ea""
								.Font.Size = fontSize
								.Font.Bold = msoFalse
								.Font.Color.ObjectThemeColor = txtColor
							End With

							' 数值格式化
							SmartNumberFormat txtRng

							' 设置底部边框
							With cell.Borders(ppBorderBottom)
								If r = rows Then
									' 尾行特殊处理
									.Weight = thick
									.ForeColor.ObjectThemeColor = bdColor1
								Else
									' 普通行处理
									.Weight = thin
									.ForeColor.ObjectThemeColor = bdColor2
								End If
							End With
						Next c
					Next r
				End Sub

				Private Sub SmartNumberFormat(rng As TextRange)
					Dim original As String
					Dim isPercentage As Boolean
					Dim numStr As String
					Dim numValue As Double
					Dim formatted As String
					Dim negativeColor As Long
					Dim pos As Integer

					' 设置负数为红色
					negativeColor = RGB(255, 0, 0)

					' 获取并清理原始文本
					original = Trim(rng.Text)
					If Len(original) = 0 Then Exit Sub

					' 检查百分比符号
					isPercentage = (Right(original, 1) = ""%"")
					If isPercentage Then
						numStr = Trim(Left(original, Len(original) - 1))
					Else
						numStr = original
					End If

					' 跨区域安全的数字解析
					If Not IsNumeric(numStr) Then Exit Sub

					' 处理不同区域设置的小数点
					If InStr(numStr, "","") > 0 And InStr(numStr, ""."") > 0 Then
						' 如果同时有逗号和点号，保留最后一个作为小数点
						If InStrRev(numStr, "","") > InStrRev(numStr, ""."") Then
							numStr = Replace(numStr, ""."", """")
							numStr = Replace(numStr, "","", ""."")
						Else
							numStr = Replace(numStr, "","", """")
						End If
					Else
						' 替换逗号为点号（欧洲格式支持）
						numStr = Replace(numStr, "","", ""."")
					End If

					' 转换数字
					numValue = CDbl(numStr)

					' 构建格式字符串
					Dim formatStr As String
					If decimalPlaces > 0 Then
						formatStr = ""#,##0."" & String(decimalPlaces, ""0"")
					Else
						formatStr = ""#,##0""
					End If

					' 应用格式化
					formatted = Format(numValue, formatStr)

					' 添加百分比符号
					If isPercentage Then
						formatted = formatted & ""%""
					End If

					' 仅当格式变化时更新文本
					If original <> formatted Then
						rng.Text = formatted
					End If

					' 设置负数颜色
					If numValue < 0 Then
						rng.Font.Color.RGB = negativeColor
					End If
				End Sub";
			ExHandler.Run(() => VbaExecutor.ExecuteVbaCode(app,vbaCode,"FormatAllTables"),"格式化表格");
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

		private static void SmartNumberFormat(NETOP.TextRange textRange,int decimalPlaces)
		{
			var negativeText = ColorTranslator.ToOle(Color.Red);
			var original = textRange.Text.Trim();
			if(string.IsNullOrEmpty(original)) return;

			var isPercentage = original.EndsWith("%");
			var numStr = isPercentage ? original.Substring(0,original.Length - 1) : original;

			if(!double.TryParse(
					numStr,
					NumberStyles.Any,
					CultureInfo.InvariantCulture,
					out var num
				))
			{
				return;
			}
			// 构造格式字符串，例如 "N0", "N2", "N3"
			var format = $"N{decimalPlaces}";
			var formatted = num.ToString(format);

			if(isPercentage)
				formatted += "%";

			if(textRange.Text == formatted) return;
			textRange.Text = formatted;

			if(num < 0)
				textRange.Font.Color.RGB = negativeText;
		}

		#endregion Private Methods
	}
}
using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using System;
using System.Collections.Generic;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 图表格式化辅助类 提供图表的格式化功能
	/// </summary>
	internal static class ChartFormatHelper
	{
		/// <summary>
		/// 格式化图表文本
		/// </summary>
		/// <param name="shape"> 包含图表的形状对象 </param>
		internal static void FormatChartText(NETOP.Shape shape)
		{
			// 参数验证
			if(shape==null||ShapeUtils.IsInvalidComObject(shape))
			{
				Profiler.LogMessage("无效的图表形状对象");
				return;
			}

			// 获取并验证图表对象
			NETOP.Chart chart;
			try
			{
				chart=shape.Chart;
				if(chart==null||ShapeUtils.IsInvalidComObject(chart))
				{
					Profiler.LogMessage("无法获取有效图表对象");
					return;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"获取图表对象时出错: {ex.Message}");
				return;
			}

			// 从配置加载参数
			var config = FormattingConfig.Instance.Chart;
			string fontFamily = config.RegularFont.Name;
			float regularSize = config.RegularFont.Size;
			float titleSize = config.TitleFont.Size;
			bool titleBold = config.TitleFont.Bold;

			// 设置图表各部分的字体
			SetChartTitleFont(chart,fontFamily,titleSize,titleBold);
			SetChartLegendFont(chart,fontFamily,regularSize);
			SetChartDataTableFont(chart,fontFamily,regularSize);

			SetChartDataLabelsFont(chart,fontFamily,regularSize);
			SetChartAxesFont(chart,regularSize);
		}

		/// <summary>
		/// 设置图表标题字体
		/// </summary>
		private static void SetChartTitleFont(NETOP.Chart chart,string fontFamily,float size,bool bold)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasTitle&&chart.ChartTitle!=null&&!ShapeUtils.IsInvalidComObject(chart.ChartTitle))
				{
					var font = chart.ChartTitle.Font;
					if(font!=null&&!ShapeUtils.IsInvalidComObject(font))
					{
						font.Name=fontFamily;
						font.Bold=bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
						font.Size=size;
					}
				}
			},"设置图表标题字体");
		}

		/// <summary>
		/// 设置图例字体
		/// </summary>
		private static void SetChartLegendFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasLegend&&chart.Legend!=null&&!ShapeUtils.IsInvalidComObject(chart.Legend))
				{
					var font = chart.Legend.Font;
					if(font!=null&&!ShapeUtils.IsInvalidComObject(font))
					{
						font.Name=fontFamily;
						font.Size=size;
					}
				}
			},"设置图例字体");
		}

		/// <summary>
		/// 设置数据表字体
		/// </summary>
		private static void SetChartDataTableFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasDataTable&&chart.DataTable!=null&&!ShapeUtils.IsInvalidComObject(chart.DataTable))
				{
					var font = chart.DataTable.Font;
					if(font!=null&&!ShapeUtils.IsInvalidComObject(font))
					{
						font.Name=fontFamily;
						font.Size=size;
					}
				}
			},"设置数据表字体");
		}

		/// <summary>
		/// 设置数据标签字体
		/// </summary>
		private static void SetChartDataLabelsFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				dynamic seriesCollection = chart.SeriesCollection();
				if(seriesCollection==null) return;

				// 使用索引访问方式，避免 NetOffice 类型转换问题
				int seriesCount = 0;
				try
				{
					seriesCount=seriesCollection.Count;
				} catch
				{
					// 如果无法获取 Count，尝试遍历
					try
					{
						foreach(dynamic series in seriesCollection)
						{
							if(series==null) continue;
							SetDataLabelsFontForSeries(series,fontFamily,size);
						}
					} catch { /* 忽略遍历错误 */ }
					return;
				}

				// 使用索引方式访问每个系列，避免类型转换异常
				for(int i = 1;i<=seriesCount;i++)
				{
					try
					{
						dynamic series = seriesCollection[i];
						if(series==null) continue;
						SetDataLabelsFontForSeries(series,fontFamily,size);
					} catch
					{
						// 继续处理下一个系列
						continue;
					}
				}
			},"设置数据标签字体");
		}

		/// <summary>
		/// 为单个系列设置数据标签字体（辅助方法）
		/// </summary>
		private static void SetDataLabelsFontForSeries(dynamic series,string fontFamily,float size)
		{
			try
			{
				// 检查系列是否有数据标签
				bool hasDataLabels = false;
				try
				{
					hasDataLabels=series.HasDataLabels;
				} catch { return; }

				if(!hasDataLabels) return;

				// 获取数据标签
				dynamic dataLabels = null;
				try
				{
					dataLabels=series.DataLabels();
				} catch { return; }

				if(dataLabels==null||ShapeUtils.IsInvalidComObject(dataLabels)) return;

				// 设置字体
				dynamic font = null;
				try
				{
					font=dataLabels.Font;
				} catch { return; }

				if(font==null||ShapeUtils.IsInvalidComObject(font)) return;

				font.Name=fontFamily;
				font.Size=size;
			} catch
			{
				// 忽略单个系列的设置错误，继续处理其他系列
			}
		}

		/// <summary>
		/// 设置坐标轴字体
		/// </summary>
		private static void SetChartAxesFont(NETOP.Chart chart,float size)
		{
			ExHandler.Run(() =>
			{
				// 检查图表类型是否支持坐标轴
				XlChartType chartType = chart.ChartType;
				var nonAxisCharts = new HashSet<XlChartType>
				{
					XlChartType.xlPie, XlChartType.xl3DPie, XlChartType.xlDoughnut,
					XlChartType.xlPieOfPie, XlChartType.xlBarOfPie,
					XlChartType.xlRadar, XlChartType.xlRadarFilled
				};

				if(nonAxisCharts.Contains(chartType))
					return;

				// 设置所有可能的坐标轴
				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlPrimary,size);
				SafeSetAxis(chart,XlAxisType.xlCategory,XlAxisGroup.xlSecondary,size);
				SafeSetAxis(chart,XlAxisType.xlValue,XlAxisGroup.xlSecondary,size);
			},"设置坐标轴字体");
		}

		/// <summary>
		/// 安全设置坐标轴字体
		/// </summary>
		private static void SafeSetAxis(NETOP.Chart chart,XlAxisType axisType,XlAxisGroup axisGroup,float size)
		{
			// 检查图表对象是否有效
			if(chart==null||ShapeUtils.IsInvalidComObject(chart))
			{
				Profiler.LogMessage($"[SafeSetAxis] 图表对象无效或已释放");
				return;
			}

			// 从配置加载字体设置
			var config = FormattingConfig.Instance.Chart;

			NETOP.Axis axis;
			// 获取坐标轴对象
			try
			{
				axis=(NETOP.Axis) chart.Axes(axisType,axisGroup);
				if(ShapeUtils.IsInvalidComObject(axis))
				{
					Profiler.LogMessage($"[SafeSetAxis] 坐标轴 {axisType}-{axisGroup} 对象无效");
					return;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"[SafeSetAxis] 获取坐标轴 {axisType}-{axisGroup} 时出错: {ex.Message}");
				return;
			}

			if(axis==null||ShapeUtils.IsInvalidComObject(axis)) return;

			// 刻度标签设置 - 添加异常处理
			try
			{
				if(axis.TickLabels!=null&&!ShapeUtils.IsInvalidComObject(axis.TickLabels))
				{
					var tickLabels = axis.TickLabels;
					if(tickLabels.Font!=null&&!ShapeUtils.IsInvalidComObject(tickLabels.Font))
					{
						tickLabels.Font.Name=config.RegularFont.Name;
						// 注意：ChartFont 不支持 NameFarEast 属性，仅设置 Name
						tickLabels.Font.Size=size;
					}
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"[SafeSetAxis] 设置刻度标签时出错: {ex.Message}");
			}

			// 坐标轴标题设置 - 添加异常处理
			try
			{
				bool hasTitle = false;
				try
				{
					hasTitle=axis.HasTitle;
				} catch { Profiler.LogMessage($"[SafeSetAxis] 无法访问坐标轴 {axisType}-{axisGroup} 的HasTitle属性"); }

				if(hasTitle&&axis.AxisTitle!=null&&!ShapeUtils.IsInvalidComObject(axis.AxisTitle))
				{
					var axisTitle = axis.AxisTitle;
					if(axisTitle.Font!=null&&!ShapeUtils.IsInvalidComObject(axisTitle.Font))
					{
						axisTitle.Font.Name=config.RegularFont.Name;
						// 注意：ChartFont 不支持 NameFarEast 属性，仅设置 Name
						axisTitle.Font.Size=size;
					}
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"[SafeSetAxis] 设置坐标轴标题时出错: {ex.Message}");
			}
		}
	}
}

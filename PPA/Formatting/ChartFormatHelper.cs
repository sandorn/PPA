using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;
using PPA.Shape;
using System;
using System.Collections.Generic;
using System.Drawing;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 图表格式化辅助类 提供图表的格式化功能
	/// </summary>
	/// <remarks>
	/// 构造函数，通过依赖注入获取配置和服务
	/// </remarks>
	/// <param name="config">格式化配置</param>
	/// <param name="shapeHelper">形状工具服务（可选）</param>
	internal class ChartFormatHelper : IChartFormatHelper
	{
		private readonly IFormattingConfig _config;
		private readonly IShapeHelper _shapeHelper;

		public ChartFormatHelper(IFormattingConfig config, IShapeHelper shapeHelper = null)
		{
			_config = config ?? throw new System.ArgumentNullException(nameof(config));

			// 如果未注入服务，尝试从 DI 容器获取
			if (shapeHelper == null)
			{
				var addIn = Globals.ThisAddIn;
				if (addIn != null && addIn.ServiceProvider != null)
				{
					shapeHelper = addIn.ServiceProvider.GetService(typeof(IShapeHelper)) as IShapeHelper;
				}
			}

			// 如果仍然为 null，创建新实例（向后兼容）
			_shapeHelper = shapeHelper ?? new ShapeUtils();
		}

		/// <summary>
		/// 格式化图表文本
		/// </summary>
		/// <param name="shape"> 包含图表的形状对象 </param>
		public void FormatChartText(NETOP.Shape shape)
		{
			// 参数验证
			if(shape==null||_shapeHelper.IsInvalidComObject(shape))
			{
				Profiler.LogMessage("无效的图表形状对象");
				return;
			}

			// 获取并验证图表对象
			NETOP.Chart chart;
			try
			{
				chart=shape.Chart;
				if(chart==null||_shapeHelper.IsInvalidComObject(chart))
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
			var config = _config.Chart;
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
		/// 格式化图表文本（抽象接口版本）
		/// </summary>
		/// <param name="shape"> 包含图表的抽象形状对象 </param>
		public void FormatChartText(IShape shape)
		{
			Profiler.LogMessage($"FormatChartText(IShape) 被调用，shape={shape?.GetType().Name ?? "null"}", "INFO");
			if(shape==null)
			{
				Profiler.LogMessage("FormatChartText: shape 为 null，返回", "WARN");
				return;
			}

			// 检查 PowerPoint 适配器
			if(shape is PowerPointShape pptShape)
			{
				Profiler.LogMessage("FormatChartText: 检测到 PowerPointShape，使用 PowerPoint 格式化", "INFO");
				FormatChartText(pptShape.NativeObject);
				return;
			}

			if(shape is IComWrapper<NETOP.Shape> typed)
			{
				Profiler.LogMessage("FormatChartText: 检测到 IComWrapper<NETOP.Shape>，使用 PowerPoint 格式化", "INFO");
				FormatChartText(typed.NativeObject);
				return;
			}

			var native = (shape as IComWrapper)?.NativeObject as NETOP.Shape;
			if(native!=null)
			{
				Profiler.LogMessage("FormatChartText: 检测到 NetOffice.Shape，使用 PowerPoint 格式化", "INFO");
				FormatChartText(native);
				return;
			}

			Profiler.LogMessage($"FormatChartText: 未知的形状类型 {shape.GetType().FullName}，无法格式化", "ERROR");
		}


		private static T SafeGet<T>(System.Func<T> getter, T @default = default)
		{
			try { return getter(); } catch { return @default; }
		}

		private static void SafeSet(System.Action action)
		{
			try { action(); } catch { }
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

		/// <summary>
		/// 设置图表标题字体
		/// </summary>
		private void SetChartTitleFont(NETOP.Chart chart,string fontFamily,float size,bool bold)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasTitle&&chart.ChartTitle!=null&&!_shapeHelper.IsInvalidComObject(chart.ChartTitle))
				{
					var font = chart.ChartTitle.Font;
					if(font!=null&&!_shapeHelper.IsInvalidComObject(font))
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
		private void SetChartLegendFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasLegend&&chart.Legend!=null&&!_shapeHelper.IsInvalidComObject(chart.Legend))
				{
					var font = chart.Legend.Font;
					if(font!=null&&!_shapeHelper.IsInvalidComObject(font))
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
		private void SetChartDataTableFont(NETOP.Chart chart,string fontFamily,float size)
		{
			ExHandler.Run(() =>
			{
				if(chart.HasDataTable&&chart.DataTable!=null&&!_shapeHelper.IsInvalidComObject(chart.DataTable))
				{
					var font = chart.DataTable.Font;
					if(font!=null&&!_shapeHelper.IsInvalidComObject(font))
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
		private void SetChartDataLabelsFont(NETOP.Chart chart,string fontFamily,float size)
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
		private void SetDataLabelsFontForSeries(dynamic series,string fontFamily,float size)
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

				if(dataLabels==null||_shapeHelper.IsInvalidComObject(dataLabels)) return;

				// 设置字体
				dynamic font = null;
				try
				{
					font=dataLabels.Font;
				} catch { return; }

				if(font==null||_shapeHelper.IsInvalidComObject(font)) return;

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
		private void SetChartAxesFont(NETOP.Chart chart,float size)
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
		private void SafeSetAxis(NETOP.Chart chart,XlAxisType axisType,XlAxisGroup axisGroup,float size)
		{
			// 检查图表对象是否有效
			if(chart==null||_shapeHelper.IsInvalidComObject(chart))
			{
				Profiler.LogMessage($"[SafeSetAxis] 图表对象无效或已释放");
				return;
			}

			// 从配置加载字体设置
			var config = _config.Chart;

			NETOP.Axis axis;
			// 获取坐标轴对象
			try
			{
				axis=(NETOP.Axis) chart.Axes(axisType,axisGroup);
				if(_shapeHelper.IsInvalidComObject(axis))
				{
					Profiler.LogMessage($"[SafeSetAxis] 坐标轴 {axisType}-{axisGroup} 对象无效");
					return;
				}
			} catch(Exception ex)
			{
				Profiler.LogMessage($"[SafeSetAxis] 获取坐标轴 {axisType}-{axisGroup} 时出错: {ex.Message}");
				return;
			}

			if(axis==null||_shapeHelper.IsInvalidComObject(axis)) return;

			// 刻度标签设置 - 添加异常处理
			try
			{
				if(axis.TickLabels!=null&&!_shapeHelper.IsInvalidComObject(axis.TickLabels))
				{
					var tickLabels = axis.TickLabels;
					if(tickLabels.Font!=null&&!_shapeHelper.IsInvalidComObject(tickLabels.Font))
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

				if(hasTitle&&axis.AxisTitle!=null&&!_shapeHelper.IsInvalidComObject(axis.AxisTitle))
				{
					var axisTitle = axis.AxisTitle;
					if(axisTitle.Font!=null&&!_shapeHelper.IsInvalidComObject(axisTitle.Font))
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

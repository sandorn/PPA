using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
    /// <summary>
    /// 格式化辅助类，提供表格、文本和图表的格式化功能。
    /// 
    /// <para>相关文档：</para>
    /// <para>- PowerPoint 主题颜色参考：<see href="../../docs/PPT主题颜色参考.md">docs/PPT主题颜色参考.md</see></para>
    /// </summary>
    public static class FormatHelper
    {
        private static readonly int NegativeTextColor = ColorTranslator.ToOle(Color.Red);

        #region Internal Methods

        internal static void ApplyTextFormatting(NETOP.Shape shp)
        {
            ExHandler.Run(() =>
            {
                // 从配置加载参数
                var config = FormattingConfig.Instance.Text;

                // 设置文本框边距（使用配置）
                var textFrame = shp.TextFrame;
                var margins = config.Margins;
                textFrame.MarginTop = ConfigHelper.CmToPoints(margins.Top);
                textFrame.MarginBottom = ConfigHelper.CmToPoints(margins.Bottom);
                textFrame.MarginLeft = ConfigHelper.CmToPoints(margins.Left);
                textFrame.MarginRight = ConfigHelper.CmToPoints(margins.Right);

                // 设置字体属性（使用配置）
                var tfont = textFrame.TextRange.Font;
                var fontConfig = config.Font;
                tfont.Name = fontConfig.Name;
                tfont.NameFarEast = fontConfig.NameFarEast;
                tfont.Color.ObjectThemeColor = ConfigHelper.GetThemeColorIndex(fontConfig.ThemeColor);
                tfont.Bold = fontConfig.Bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                tfont.Size = fontConfig.Size;

                // 设置段落格式（使用配置）
                var paragraph = textFrame.TextRange.ParagraphFormat;
                var paraConfig = config.Paragraph;
                paragraph.FarEastLineBreakControl = paraConfig.FarEastLineBreakControl ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                paragraph.HangingPunctuation = paraConfig.HangingPunctuation ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                paragraph.BaseLineAlignment = NETOP.Enums.PpBaselineAlignment.ppBaselineAlignAuto;
                paragraph.Alignment = ConfigHelper.GetParagraphAlignment(paraConfig.Alignment);
                paragraph.WordWrap = paraConfig.WordWrap ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                paragraph.SpaceBefore = paraConfig.SpaceBefore;
                paragraph.SpaceAfter = paraConfig.SpaceAfter;
                paragraph.SpaceWithin = paraConfig.SpaceWithin;

                // 设置项目符号（使用配置）
                var bullet = paragraph.Bullet;
                var bulletConfig = config.Bullet;
                bullet.Type = ConfigHelper.GetBulletType(bulletConfig.Type);
                bullet.Character = bulletConfig.Character;
                bullet.Font.Name = bulletConfig.FontName;
                bullet.RelativeSize = bulletConfig.RelativeSize;
                bullet.Font.Color.ObjectThemeColor = ConfigHelper.GetThemeColorIndex(bulletConfig.ThemeColor);

                // 设置悬挂缩进（使用配置）
                textFrame.Ruler.Levels[1].LeftMargin = ConfigHelper.CmToPoints(config.LeftIndent);
            });
        }

        internal static void FormatChartText(NETOP.Shape shape)
        {
            // 参数验证
            if (shape == null || ShapeUtils.IsInvalidComObject(shape))
            {
                Profiler.LogMessage("无效的图表形状对象");
                return;
            }

            // 获取并验证图表对象
            NETOP.Chart chart;
            try
            {
                chart = shape.Chart;
                if (chart == null || ShapeUtils.IsInvalidComObject(chart))
                {
                    Profiler.LogMessage("无法获取有效图表对象");
                    return;
                }
            }
            catch (Exception ex)
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
            SetChartTitleFont(chart, fontFamily, titleSize, titleBold);
            SetChartLegendFont(chart, fontFamily, regularSize);
            SetChartDataTableFont(chart, fontFamily, regularSize);

            SetChartDataLabelsFont(chart, fontFamily, regularSize);
            SetChartAxesFont(chart, regularSize);
        }

        #region Private Chart Formatting Methods

        /// <summary>
        /// 设置图表标题字体
        /// </summary>
        private static void SetChartTitleFont(NETOP.Chart chart, string fontFamily, float size, bool bold)
        {
            ExHandler.Run(() =>
            {
                if (chart.HasTitle && chart.ChartTitle != null && !ShapeUtils.IsInvalidComObject(chart.ChartTitle))
                {
                    var font = chart.ChartTitle.Font;
                    if (font != null && !ShapeUtils.IsInvalidComObject(font))
                    {
                        font.Name = fontFamily;
                        font.Bold = bold ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                        font.Size = size;
                    }
                }
            }, "设置图表标题字体");
        }

        /// <summary>
        /// 设置图例字体
        /// </summary>
        private static void SetChartLegendFont(NETOP.Chart chart, string fontFamily, float size)
        {
            ExHandler.Run(() =>
            {
                if (chart.HasLegend && chart.Legend != null && !ShapeUtils.IsInvalidComObject(chart.Legend))
                {
                    var font = chart.Legend.Font;
                    if (font != null && !ShapeUtils.IsInvalidComObject(font))
                    {
                        font.Name = fontFamily;
                        font.Size = size;
                    }
                }
            }, "设置图例字体");
        }

        /// <summary>
        /// 设置数据表字体
        /// </summary>
        private static void SetChartDataTableFont(NETOP.Chart chart, string fontFamily, float size)
        {
            ExHandler.Run(() =>
            {
                if (chart.HasDataTable && chart.DataTable != null && !ShapeUtils.IsInvalidComObject(chart.DataTable))
                {
                    var font = chart.DataTable.Font;
                    if (font != null && !ShapeUtils.IsInvalidComObject(font))
                    {
                        font.Name = fontFamily;
                        font.Size = size;
                    }
                }
            }, "设置数据表字体");
        }

        /// <summary>
        /// 设置数据标签字体
        /// </summary>
        private static void SetChartDataLabelsFont(NETOP.Chart chart, string fontFamily, float size)
        {
            ExHandler.Run(() =>
            {
                dynamic seriesCollection = chart.SeriesCollection();
                if (seriesCollection == null) return;

                // 使用索引访问方式，避免 NetOffice 类型转换问题
                int seriesCount = 0;
                try
                {
                    seriesCount = seriesCollection.Count;
                }
                catch
                {
                    // 如果无法获取 Count，尝试遍历
                    try
                    {
                        foreach (dynamic series in seriesCollection)
                        {
                            if (series == null) continue;
                            SetDataLabelsFontForSeries(series, fontFamily, size);
                        }
                    }
                    catch { /* 忽略遍历错误 */ }
                    return;
                }

                // 使用索引方式访问每个系列，避免类型转换异常
                for (int i = 1; i <= seriesCount; i++)
                {
                    try
                    {
                        dynamic series = seriesCollection[i];
                        if (series == null) continue;
                        SetDataLabelsFontForSeries(series, fontFamily, size);
                    }
                    catch
                    {
                        // 继续处理下一个系列
                        continue;
                    }
                }
            }, "设置数据标签字体");
        }

        /// <summary>
        /// 为单个系列设置数据标签字体（辅助方法）
        /// </summary>
        private static void SetDataLabelsFontForSeries(dynamic series, string fontFamily, float size)
        {
            try
            {
                // 检查系列是否有数据标签
                bool hasDataLabels = false;
                try
                {
                    hasDataLabels = series.HasDataLabels;
                }
                catch { return; }

                if (!hasDataLabels) return;

                // 获取数据标签
                dynamic dataLabels = null;
                try
                {
                    dataLabels = series.DataLabels();
                }
                catch { return; }

                if (dataLabels == null || ShapeUtils.IsInvalidComObject(dataLabels)) return;

                // 设置字体
                dynamic font = null;
                try
                {
                    font = dataLabels.Font;
                }
                catch { return; }

                if (font == null || ShapeUtils.IsInvalidComObject(font)) return;

                font.Name = fontFamily;
                font.Size = size;
            }
            catch
            {
                // 忽略单个系列的设置错误，继续处理其他系列
            }
        }

        /// <summary>
        /// 设置坐标轴字体
        /// </summary>
        private static void SetChartAxesFont(NETOP.Chart chart, float size)
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

                if (nonAxisCharts.Contains(chartType))
                    return;

                // 设置所有可能的坐标轴
                SafeSetAxis(chart, XlAxisType.xlCategory, XlAxisGroup.xlPrimary, size);
                SafeSetAxis(chart, XlAxisType.xlValue, XlAxisGroup.xlPrimary, size);
                SafeSetAxis(chart, XlAxisType.xlCategory, XlAxisGroup.xlSecondary, size);
                SafeSetAxis(chart, XlAxisType.xlValue, XlAxisGroup.xlSecondary, size);
            }, "设置坐标轴字体");
        }

        #endregion Private Chart Formatting Methods

        /// <summary>
        /// 对表格进行高性能格式化。
        /// </summary>
        /// <param name="tbl">要格式化的 PowerPoint 表格对象。</param>
        internal static void FormatTables(NETOP.Table tbl)
        {
            // 从配置加载参数
            var config = FormattingConfig.Instance.Table;
            var tableConfig = config;

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
            tbl.ApplyStyle(styleId, false);
            var tableSettings = tableConfig.TableSettings;
            tbl.FirstRow = tableSettings.FirstRow;
            tbl.FirstCol = tableSettings.FirstCol;
            tbl.LastRow = tableSettings.LastRow;
            tbl.LastCol = tableSettings.LastCol;
            tbl.HorizBanding = tableSettings.HorizBanding;
            tbl.VertBanding = tableSettings.VertBanding;

            // --- 3. 性能优化：批处理模式 --- 
            // 预先创建批处理集合
            var firstRowCells = new List<NETOP.Cell>();
            var lastRowCells = new List<NETOP.Cell>();
            var dataRowCells = new List<NETOP.Cell>();


            // 第一步：收集所有单元格到不同集合
            for (int r = 1; r <= rows; r++)
            {
                var row = tbl.Rows[r];
                for (int c = 1; c <= cols; c++)
                {
                    var cell = row.Cells[c];
                    dataRowCells.Add(cell);
                    // 只收集引用，不立即处理
                    if (r == 1)
                        firstRowCells.Add(cell);
                    else if (r == rows)
                        lastRowCells.Add(cell);
                }
            }

            //批量处理数据行
            FormatDataRowCells(dataRowCells, fontName, fontNameFarEast, fontSize, dk1, thin, a2, useAutoNum, decimalPlacesValue, tableConfig.NegativeTextColor);

            //批量处理标题行和尾行
            FormatOutsideRowCells(firstRowCells, lastRowCells, 
                tableConfig.HeaderRowFont.Name, 
                tableConfig.HeaderRowFont.NameFarEast, 
                bigFontSize, 
                dk1, 
                thick, 
                a1);

        }

        /// <summary>
        /// 批量处理首末行的单元格，减少重复操作和COM调用
        /// </summary>
        private static void FormatOutsideRowCells(List<NETOP.Cell> firstRowCells, List<NETOP.Cell> lastRowCells, string fontName, string fontNameFarEast, float fontSize, MsoThemeColorIndex txtColor, float borderWidth, MsoThemeColorIndex borderColor)
        {
			// 设置首行上下边框
			for(int i = 0; i < firstRowCells.Count; i++)
            {
                var cell = firstRowCells[i];
                cell.Shape.Fill.Visible = MsoTriState.msoFalse;
                var textRange = cell.Shape.TextFrame.TextRange;
                SetFontProperties(textRange, fontName, fontNameFarEast, fontSize, MsoTriState.msoTrue, txtColor);

                textRange.ParagraphFormat.Alignment = NETOP.Enums.PpParagraphAlignment.ppAlignCenter;

                // 边框
                SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderTop, borderWidth, (object)borderColor);
                SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderBottom, borderWidth, (object)borderColor);
            }

			// 设置末行下边框
			for(int i = 0; i < lastRowCells.Count; i++)
            {
                var cell = lastRowCells[i];
                SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderBottom, borderWidth, (object)borderColor);
            }
        }

        /// <summary>
        /// 批量处理数据行的单元格，使用更高效的处理方式
        /// </summary>
        private static void FormatDataRowCells(List<NETOP.Cell> cells, string fontName, string fontNameFarEast, float fontSize, MsoThemeColorIndex txtColor, float thinBorderWidth, MsoThemeColorIndex thinBorderColor, bool autonum, int decimalPlaces, int negativeTextColor)
        {
            int cellCount = cells.Count;

            for (int i = 0; i < cellCount; i++)
            {
                var cell = cells[i];
                cell.Shape.Fill.Visible = MsoTriState.msoFalse;

                var textRange = cell.Shape.TextFrame.TextRange;
                SetFontProperties(textRange, fontName, fontNameFarEast, fontSize, MsoTriState.msoFalse, txtColor);

                // 智能优化：只对非空文本进行数字格式化
                if (autonum && !string.IsNullOrEmpty(textRange.Text.Trim()))
                {
                    SmartNumberFormat(textRange, decimalPlaces, negativeTextColor);
                }

                SetBorder(cell, NETOP.Enums.PpBorderType.ppBorderTop, thinBorderWidth, (object)thinBorderColor, 0.5f);
            }
        }


        /// <summary>
        /// 批量设置字体属性，减少 COM 调用次数。
        /// </summary>
        private static void SetFontProperties(NETOP.TextRange textRange, string name, string nameFarEast, float size, MsoTriState bold, MsoThemeColorIndex color)
        {
            // 关键：通过 .Font 来访问字体属性
            textRange.Font.Name = name;
            textRange.Font.NameFarEast = nameFarEast;
            textRange.Font.Size = size;
            textRange.Font.Bold = bold;
            textRange.Font.Color.ObjectThemeColor = color;
        }

        /// <summary>
        /// 高性能数字格式化，针对大量单元格优化，在必要时修改文本和颜色
        /// </summary>
        private static void SmartNumberFormat(NETOP.TextRange textRange, int decimalPlaces, int negativeTextColor)
        {
            // 性能优化1: 直接访问文本，避免多次Trim操作
            string text = textRange.Text;
            if (string.IsNullOrEmpty(text)) return;

            // 预先计算可能的百分比符号位置
            int length = text.Length;
            bool isPercentage = length > 0 && text[length - 1] == '%';

            // 获取需要解析的数字部分
            string numStr = isPercentage ? text.Substring(0, length - 1).Trim() : text.Trim();

            // 性能优化2: 快速检查是否可能是数字
            if (string.IsNullOrEmpty(numStr) ||
                (!char.IsDigit(numStr[0]) && numStr[0] != '-' && numStr[0] != '.' && numStr[0] != '+'))
            {
                return;
            }

            // 性能优化3: 尝试解析数字
            if (!double.TryParse(numStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
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
            if (isPercentage)
            {
                formatted += "%";
            }

            // 性能优化5: 避免不必要的COM调用 - 只有当文本真的需要改变时才设置
            if (text != formatted)
            {
                textRange.Text = formatted;
            }

            // 性能优化6: 负数颜色设置 - 只在需要时调用
            if (num < 0)
            {
                textRange.Font.Color.RGB = negativeTextColor;
            }
        }

        #endregion Internal Methods

        #region Private Methods

        private static void SafeSetAxis(NETOP.Chart chart, XlAxisType axisType, XlAxisGroup axisGroup, float size)
        {
            // 检查图表对象是否有效
            if (chart == null || ShapeUtils.IsInvalidComObject(chart))
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
                axis = (NETOP.Axis)chart.Axes(axisType, axisGroup);
                if (ShapeUtils.IsInvalidComObject(axis))
                {
                    Profiler.LogMessage($"[SafeSetAxis] 坐标轴 {axisType}-{axisGroup} 对象无效");
                    return;
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"[SafeSetAxis] 获取坐标轴 {axisType}-{axisGroup} 时出错: {ex.Message}");
                return;
            }

            if (axis == null || ShapeUtils.IsInvalidComObject(axis)) return;

            // 刻度标签设置 - 添加异常处理
            try
            {
                if (axis.TickLabels != null && !ShapeUtils.IsInvalidComObject(axis.TickLabels))
                {
                    var tickLabels = axis.TickLabels;
                    if (tickLabels.Font != null && !ShapeUtils.IsInvalidComObject(tickLabels.Font))
                    {
                        tickLabels.Font.Name = config.RegularFont.Name;
                        // 注意：ChartFont 不支持 NameFarEast 属性，仅设置 Name
                        tickLabels.Font.Size = size;
                    }
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"[SafeSetAxis] 设置刻度标签时出错: {ex.Message}");
            }

            // 坐标轴标题设置 - 添加异常处理
            try
            {
                bool hasTitle = false;
                try
                {
                    hasTitle = axis.HasTitle;
                }
                catch { Profiler.LogMessage($"[SafeSetAxis] 无法访问坐标轴 {axisType}-{axisGroup} 的HasTitle属性"); }

                if (hasTitle && axis.AxisTitle != null && !ShapeUtils.IsInvalidComObject(axis.AxisTitle))
                {
                    var axisTitle = axis.AxisTitle;
                    if (axisTitle.Font != null && !ShapeUtils.IsInvalidComObject(axisTitle.Font))
                    {
                        axisTitle.Font.Name = config.RegularFont.Name;
                        // 注意：ChartFont 不支持 NameFarEast 属性，仅设置 Name
                        axisTitle.Font.Size = size;
                    }
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"[SafeSetAxis] 设置坐标轴标题时出错: {ex.Message}");
            }
        }

        // 判断 tcolor 类型
        private static void SetBorder(NETOP.Cell cell, NETOP.Enums.PpBorderType borderType, float setWeight, object tcolor, float transparency = 0)
        {
            var border = cell.Borders[borderType];

            // weight 为 0,隐藏条线
            if (setWeight <= 0f)
            {
                border.Weight = setWeight;
                border.Visible = MsoTriState.msoFalse;
            }
            else
            {
                border.Weight = setWeight;
                border.Visible = MsoTriState.msoTrue;
                border.Transparency = transparency;
                // 使用模式匹配简化颜色逻辑
                if (tcolor is MsoThemeColorIndex themeColor) border.ForeColor.ObjectThemeColor = themeColor;
                else if (tcolor is int rgbColor) border.ForeColor.RGB = rgbColor;
                else border.ForeColor.RGB = 0; // 默认黑色
            }
        }

        #endregion Private Methods
    }
}
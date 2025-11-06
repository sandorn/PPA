using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using ToastAPI;
using VBAApi;
using MSOP = Microsoft.Office.Interop.PowerPoint;
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
                const MsoThemeColorIndex dk1 = MsoThemeColorIndex.msoThemeColorDark1;
                const MsoThemeColorIndex a2 = MsoThemeColorIndex.msoThemeColorAccent2;

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
                tfont.Color.ObjectThemeColor = a2;
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
                bullet.Character = 9632; // 实心方块
                bullet.Font.Name = "Arial";
                bullet.RelativeSize = 1.0f;
                bullet.Font.Color.ObjectThemeColor = dk1;

                // 设置悬挂缩进（通过 Ruler 对象）
                textFrame.Ruler.Levels[1].LeftMargin = 1.0f * 28.35f; // 厘米转磅,段落左缩进
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
            NETOP.Chart chart = null;
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

            // 定义常量
            const string fontFamily = "+mn-lt";
            const float regularSize = 8f;
            const float titleSize = 11f;

            // 1. 设置图表标题字体
            ExHandler.Run(() =>
            {
                if (chart.HasTitle && chart.ChartTitle != null && !ShapeUtils.IsInvalidComObject(chart.ChartTitle))
                {
                    if (chart.ChartTitle.Font != null && !ShapeUtils.IsInvalidComObject(chart.ChartTitle.Font))
                    {
                        chart.ChartTitle.Font.Name = fontFamily;
                        chart.ChartTitle.Font.Bold = MsoTriState.msoTrue;
                        chart.ChartTitle.Font.Size = titleSize;
                    }
                }
            }, "设置图表标题字体");

            // 2. 设置图例字体
            ExHandler.Run(() =>
            {
                if (chart.HasLegend && chart.Legend != null && !ShapeUtils.IsInvalidComObject(chart.Legend))
                {
                    if (chart.Legend.Font != null && !ShapeUtils.IsInvalidComObject(chart.Legend.Font))
                    {
                        chart.Legend.Font.Name = fontFamily;
                        chart.Legend.Font.Size = regularSize;
                    }
                }
            }, "设置图例字体");

            // 3. 设置数据表字体
            ExHandler.Run(() =>
            {
                if (chart.HasDataTable && chart.DataTable != null && !ShapeUtils.IsInvalidComObject(chart.DataTable))
                {
                    if (chart.DataTable.Font != null && !ShapeUtils.IsInvalidComObject(chart.DataTable.Font))
                    {
                        chart.DataTable.Font.Name = fontFamily;
                        chart.DataTable.Font.Size = regularSize;
                    }
                }
            }, "设置数据表字体");

            // 4. 设置数据标签字体 - 添加异常处理提高兼容性
            ExHandler.Run(() =>
            {
                try
                {
                    dynamic seriesCollection = chart.SeriesCollection();
                    if (seriesCollection == null) return;

                    foreach (dynamic series in seriesCollection)
                    {
                        if (series == null) continue;

                        try
                        {
                            // 检查是否有数据标签
                            bool hasDataLabels = false;
                            try
                            {
                                hasDataLabels = series.HasDataLabels;
                            }
                            catch { continue; }

                            if (hasDataLabels)
                            {
                                try
                                {
                                    dynamic dataLabels = series.DataLabels();
                                    if (dataLabels != null && !ShapeUtils.IsInvalidComObject(dataLabels))
                                    {
                                        if (dataLabels.Font != null && !ShapeUtils.IsInvalidComObject(dataLabels.Font))
                                        {
                                            dataLabels.Font.Name = fontFamily;
                                            dataLabels.Font.Size = regularSize;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Profiler.LogMessage($"设置数据标签时出错: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Profiler.LogMessage($"处理数据系列时出错: {ex.Message}");
                            continue; // 继续处理下一个系列
                        }
                    }
                }
                catch (Exception ex)
                {
                    Profiler.LogMessage($"获取数据系列集合时出错: {ex.Message}");
                }
            }, "设置数据标签字体");

            // 5. 设置坐标轴字体 - 增强健壮性
            ExHandler.Run(() =>
            {
                try
                {
                    // 获取图表类型并检查是否为不支持坐标轴的类型
                    XlChartType chartType;
                    try
                    {
                        chartType = chart.ChartType;
                    }
                    catch (Exception ex)
                    {
                        Profiler.LogMessage($"获取图表类型时出错: {ex.Message}");
                        return;
                    }

                    // 预定义不支持坐标轴的图表类型
                    var nonAxisCharts = new HashSet<XlChartType>
                    {
                        XlChartType.xlPie,
                        XlChartType.xl3DPie,
                        XlChartType.xlDoughnut,
                        XlChartType.xlPieOfPie,
                        XlChartType.xlBarOfPie,
                        XlChartType.xlRadar,
                        XlChartType.xlRadarFilled
                    };

                    // 对于不支持坐标轴的图表类型，跳过坐标轴设置
                    if (nonAxisCharts.Contains(chartType))
                    {
                        Profiler.LogMessage($"图表类型 {chartType} 不支持坐标轴设置，已跳过");
                        return;
                    }

                    // 逐个设置坐标轴，即使某个坐标轴设置失败也不影响其他坐标轴
                    // 主分类轴
                    SafeSetAxis(chart, XlAxisType.xlCategory, XlAxisGroup.xlPrimary, regularSize);

                    // 主值轴
                    SafeSetAxis(chart, XlAxisType.xlValue, XlAxisGroup.xlPrimary, regularSize);

                    // 次分类轴（可能不存在）
                    SafeSetAxis(chart, XlAxisType.xlCategory, XlAxisGroup.xlSecondary, regularSize);

                    // 次值轴（可能不存在）
                    SafeSetAxis(chart, XlAxisType.xlValue, XlAxisGroup.xlSecondary, regularSize);
                }
                catch (Exception ex)
                {
                    Profiler.LogMessage($"处理坐标轴设置时出错: {ex.Message}");
                }
            }, "设置坐标轴字体");
        }

        /// <summary>
        /// 对表格进行高性能格式化。
        /// </summary>
        /// <param name="tbl">要格式化的 PowerPoint 表格对象。</param>
        /// <param name="autonum">是否自动格式化数字。</param>
        /// <param name="decimalPlaces">保留的小数位数。</param>
        internal static void FormatTables(NETOP.Table tbl, bool autonum = true, int decimalPlaces = 2)
        {
            // --- 1. 预定义所有常量 ---
            const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
            const MsoThemeColorIndex dk1 = MsoThemeColorIndex.msoThemeColorDark1;
            const MsoThemeColorIndex a1 = MsoThemeColorIndex.msoThemeColorAccent1;
            const MsoThemeColorIndex a2 = MsoThemeColorIndex.msoThemeColorAccent2;

            const float thin = 1.0f, thick = 1.75f;
            const float fontSize = 9.0f, bigFontSize = 10.0f;
            const string fontName = "+mn-lt";
            const string fontNameFarEast = "+mn-ea";

            int rows = tbl.Rows.Count;
            int cols = tbl.Columns.Count;

            // --- 2. 一次性设置表格全局样式 ---
            tbl.ApplyStyle(styleId, false);
            tbl.FirstRow = true;
            tbl.FirstCol = false;
            tbl.LastRow = false;
            tbl.LastCol = false;
            tbl.HorizBanding = false;
            tbl.VertBanding = false;

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
            FormatDataRowCells(dataRowCells, fontName, fontNameFarEast, fontSize, dk1, thin, a2, autonum, decimalPlaces, NegativeTextColor);

            //批量处理标题行和尾行
            FormatOutsideRowCells(firstRowCells, lastRowCells, fontName, fontNameFarEast, bigFontSize, dk1, thick, a1);

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

        internal static void FormatTablesbyVBA(NETOP.Application app, NETOP.Slide slide)
        {
            if (app == null || slide == null)
            {
                Toast.Show("无效的应用程序或幻灯片对象", Toast.ToastType.Error);
                return;
            }

            const string styleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
            int tableCount = 0;

            // 先应用表格样式，并统计有多少个表格
            foreach (NETOP.Shape shape in slide.Shapes)
            {
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    NETOP.Table tbl = shape.Table;
                    tbl.FirstRow = true;    // 标题行（首行特殊格式）
                    tbl.FirstCol = false;   // 标题列（首列特殊格式）
                    tbl.LastRow = false;    // 汇总行（末行特殊格式）
                    tbl.LastCol = false;    // 汇总列（末列特殊格式）
                    tbl.HorizBanding = false;   //镶边行
                    tbl.VertBanding = false;    //镶边列
                    tbl.ApplyStyle(styleId, false);
                    tableCount++;
                }
            }

            if (tableCount == 0)
            {
                Toast.Show("当前幻灯片上没有表格", Toast.ToastType.Info);
                return;
            }

            // 显示等待光标
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

            // 调用新的管理器，它会自动处理模块初始化
            VbaManager.RunMacro(app, "FormatAllTables");
            Toast.Show($"成功格式化了 {tableCount} 个表格", Toast.ToastType.Success);

            // 恢复光标
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
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
                        tickLabels.Font.Name = "+mn-lt";
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
                        axisTitle.Font.Name = "+mn-lt";
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

/* 
# PPT 12 种主体颜色与 MsoThemeColorIndex 的对应关系

    中文名称 (UI 显示)	英文名称 (设计理念)	对应的 MsoThemeColorIndex 枚举值	说明
    文字/背景 - 深色 1	Dark 1  暗 1	msoThemeColorDark1	通常用于主要文本或深色背景。
    文字/背景 - 浅色 1	Light 1  光 1	msoThemeColorLight1	通常用于幻灯片背景或浅色文本。
    文字/背景 - 深色 2	Dark 2  暗 2	msoThemeColorDark2	辅助深色，用于次要文本或背景。
    文字/背景 - 浅色 2	Light 2  光 2	msoThemeColorLight2	辅助浅色，用于填充或高亮背景。
    着色 1	Accent 1  强调 1	msoThemeColorAccent1	主要强调色，通常是最突出的品牌色。
    着色 2	Accent 2  重音 2	msoThemeColorAccent2	次要强调色。
    着色 3	Accent 3  重音 3	msoThemeColorAccent3	第三强调色。
    着色 4	Accent 4  重音 4	msoThemeColorAccent4	第四强调色。
    着色 5	Accent 5  重音 5	msoThemeColorAccent5	第五强调色。
    着色 6	Accent 6  重音 6	msoThemeColorAccent6	第六强调色。
    超链接	Hyperlink  超链接	msoThemeColorHyperlink	用于未点击的超链接。
    已访问的超链接	Followed Hyperlink  点击超链接	msoThemeColorFollowedHyperlink	用于已点击的超链接。
    msoThemeColorText1 和 msoThemeColorBackground1 这两个枚举值也存在，它们在内部通常分别指向 msoThemeColorDark1 和 msoThemeColorLight1

## 主题颜色变体 (Tint and Shade)
    获取形状的填充颜色格式对象
    var fillFormat = shape.Fill.ForeColor;

    设置颜色为“着色 1”
    fillFormat.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorAccent1;

    将其调亮 40% (Tint)
    fillFormat.TintAndShade = 0.4f;

    将其调暗 20% (Shade)
    fillFormat.TintAndShade = -0.2f;

## 主题字体
    获取当前幻灯片的主主题
    var theme = slide.Design.Theme;

    获取主题字体方案
    var fontScheme = theme.ThemeFontScheme;

    获取主要字体（用于标题）
    var majorFont = fontScheme.MajorFont; // 这是一个 Font 对象，包含 .Name, .NameFarEast, .NameAscii 等属性
    Profiler.LogMessage($"标题字体: {majorFont.Name}");

    获取次要字体（用于正文）
    var minorFont = fontScheme.MinorFont;
    Profiler.LogMessage($"正文字体: {minorFont.Name}");

    应用主题字体到文本框
    shape.TextFrame.TextRange.Font.Name = fontScheme.MajorFont.Name;

    应用次要字体到文本框
    shape.TextFrame.TextRange.Font.Name = fontScheme.MinorFont.Name;
    
    主题字体别名
    var tfont = textFrame.TextRange.Font;
    tfont.Name = "+mn-lt";      // 拉丁字母使用主题的“次要字体”
    tfont.NameFarEast = "+mn-ea"; // 东亚字符（如中文）使用主题的“次要字体”

    代号	含义	对应主题中的角色	常用场景
    +mj-lt	Major Latin  主要拉丁语	主要字体 (拉丁)	标题、页眉中的西文
    +mj-ea	Major East Asian  主要东亚	主要字体 (东亚)	标题、页眉中的中文
    +mn-lt	Minor Latin  次要拉丁语	次要字体 (拉丁)	正文、备注中的西文
    +mn-ea	Minor East Asian  次要东亚	次要字体 (东亚)	正文、备注中的中文

## 主题效果
    应用一个预设的阴影效果，这个效果会与主题颜色协调
    shape.Shadow.Type = MsoShadowType.msoShadow21;

    应用一个预设的柔化边缘效果
    shape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType1;

## 整个主题颜色方案
    获取当前幻灯片的设计主题
    var theme = slide.Design.Theme;

    获取主题颜色方案
    var colorScheme = theme.ThemeColorScheme;

    获取“着色 1”的 RGB 值
    注意：GetColor 返回的是一个 MsoRGBType，可以转换为 int
    int accent1Rgb = colorScheme.GetColor(MsoThemeColorIndex.msoThemeColorAccent1);

    可以修改颜色方案（会影响到整个使用该主题的幻灯片）
    colorScheme.Colors(MsoThemeColorIndex.msoThemeColorAccent1).RGB = RGB(255, 0, 0); // 将着色1改为红色

## 总结
    功能	对象/属性	示例用途
    基础颜色	ObjectThemeColor	将形状或文本设置为 msoThemeColorAccent1
    颜色变体	TintAndShade	创建更亮或更暗的强调色
    主题字体	ThemeFontScheme	获取或应用标题/正文字体
    主题效果	EffectFormat	应用与主题协调的阴影、发光等
    颜色方案	ThemeColorScheme	获取主题中任意颜色的 RGB 值，或修改整个方案
*/
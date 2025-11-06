using NetOffice.OfficeApi.Enums;
using Project.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ToastAPI;
using NETOP = NetOffice.PowerPointApi;
using ComListExtensions;
//using NetOffice.PowerPointApi.Enums;

namespace PPA.Helpers
{
    public static class BatchHelper
    {
        #region Public Enums

        public enum AlignmentType
        {
            Left, Right, Top, Bottom, Centers, Middles, Horizontally, Vertically
        }

        #endregion Public Enums

        #region Private Methods

        /// <summary>
        /// 安全获取当前幻灯片：通过 Interop 读取 SlideIndex，再通过 NetOffice 获取，避免直接访问 View.Slide 导致的本地化类名包装失败
        /// </summary>
        private static NETOP.Slide TryGetCurrentSlide(NETOP.Application app)
        {
            if (app == null) return null;
            try
            {
                // 优先通过 Interop 取索引，避免 NetOffice 包装本地化类名
                var underlying = (app as NetOffice.ICOMObject)?.UnderlyingObject as Microsoft.Office.Interop.PowerPoint.Application;
                int slideIndex = 0;
                try { slideIndex = underlying?.ActiveWindow?.View?.Slide?.SlideIndex ?? 0; }
                catch (Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide interop读取异常: {ex.Message}"); }

                if (slideIndex > 0)
                {
                    try { return app?.ActivePresentation?.Slides[slideIndex]; }
                    catch (Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide netoffice索引获取异常: {ex.Message}"); }
                }

                // 备选1：Selection.SlideRange
                try
                {
                    var sel = app?.ActiveWindow?.Selection;
                    var sr = sel?.SlideRange;
                    if (sr != null && sr.Count >= 1)
                    {
                        try { return sr[1]; }
                        finally { sr?.Dispose(); }
                    }
                }
                catch (Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide选择范围异常: {ex.Message}"); }
            }
            catch (Exception ex) { Profiler.LogMessage($"TryGetCurrentSlide异常: {ex.Message}"); }
            return null;
        }

        #endregion Private Methods

        #region Public Methods

        public static void Bt501_Click(NETOP.Application app)
        {
            app.StartNewUndoEntry();
            var slide = TryGetCurrentSlide(app);

            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                if (sel != null)
                {
                    // 性能优化：批量收集并处理表格
                    int tableCount = 0;
                    // 一次性遍历所有选中形状，减少COM对象创建
                    if (sel is NETOP.Shape shape)
                    {
                        if (shape.HasTable == MsoTriState.msoTrue)
                        {
                            // 使用优化后的FormatTables方法，添加参数控制数字格式化
                            FormatHelper.FormatTables(shape.Table, autonum: true, decimalPlaces: 0);
                            tableCount++;
                        }
                    }
                    else if (sel is NETOP.ShapeRange shapes)
                    {
                        foreach (NETOP.Shape s in shapes)
                        {
                            if (s.HasTable == MsoTriState.msoTrue)
                            {
                                // 使用优化后的FormatTables方法，添加参数控制数字格式化
                                FormatHelper.FormatTables(s.Table, autonum: true, decimalPlaces: 0);
                                tableCount++;
                            }
                        }
                    }
                    if (tableCount > 0)
                        Toast.Show($"成功格式化 {tableCount} 个表格", Toast.ToastType.Success);
                    else
                        Toast.Show("未找到表格", Toast.ToastType.Info);
                }
                else
                {
                    // 未选中对象的情况，VBA调用
                    FormatHelper.FormatTablesbyVBA(app, slide);
                }
            }, enableTiming: true);
        }

        public static void Bt502_Click(NETOP.Application app)
        {
            app.StartNewUndoEntry();

            ExHandler.Run(() =>
            {
                // 获取选中的形状
                var selection = ShapeUtils.ValidateSelection(app);

                // 如果没有选中对象，显示提示并返回
                if (selection == null)
                {
                    Toast.Show("请先选择要格式化的形状", Toast.ToastType.Warning);
                    return;
                }

                bool hasFormatted = false;

                // 处理单个形状的情况
                if (selection is NETOP.Shape shape)
                {
                    if (shape.TextFrame?.HasText == MsoTriState.msoTrue)
                    {
                        FormatHelper.ApplyTextFormatting(shape);
                        hasFormatted = true;
                    }
                }
                // 处理多个形状的情况
                else if (selection is NETOP.ShapeRange shapeRange)
                {
                    foreach (NETOP.Shape s in shapeRange)
                    {
                        if (s.TextFrame?.HasText == MsoTriState.msoTrue)
                        {
                            FormatHelper.ApplyTextFormatting(s);
                            hasFormatted = true;
                        }
                    }
                }

                // 如果成功格式化了文本，显示成功提示
                if (hasFormatted)
                {
                    Toast.Show("格式化文本完成", Toast.ToastType.Success);
                }
                else
                {
                    Toast.Show("选中的形状中没有文本框", Toast.ToastType.Warning);
                }
            });
        }

        public static void Bt503_Click(NETOP.Application app)
        {
            app.StartNewUndoEntry();

            ExHandler.Run(() =>
            {
                var slide = TryGetCurrentSlide(app);
                if (slide == null) return;

                // 有选中对象，则处理选中的对象
                var sel = ShapeUtils.ValidateSelection(app);
                if (sel != null)
                {
                    if (sel is NETOP.Shape shape)
                    {
                        if (shape.HasChart == MsoTriState.msoTrue) FormatHelper.FormatChartText(shape);
                    }
                    else if (sel is NETOP.ShapeRange shapes)
                    {
                        foreach (NETOP.Shape s in shapes)
                            if (s.HasChart == MsoTriState.msoTrue) FormatHelper.FormatChartText(s);
                    }
                    Toast.Show("格式化图表完成", Toast.ToastType.Success);
                }
                else
                {
                    // 没有选中对象时，处理当前幻灯片上所有对象
                    foreach (NETOP.Shape shape in slide.Shapes)
                        if (shape.HasChart == MsoTriState.msoTrue) FormatHelper.FormatChartText(shape);
                    Toast.Show("格式化图表完成", Toast.ToastType.Success);
                }
            });
        }

        /// <summary>
        /// 根据选中对象创建矩形外框：
        /// 1. 选中形状时：为每个形状创建包围框并考虑边框宽度
        /// 2. 选中幻灯片时：在每个幻灯片创建页面大小的矩形
        /// 3. 无选中时：在当前幻灯片创建页面大小的矩形
        /// </summary>
        public static void Bt601_Click(NETOP.Application app)
        {
            app.StartNewUndoEntry();

            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                var currentSlide = TryGetCurrentSlide(app);

                if (currentSlide == null)
                {
                    Toast.Show("未找到当前幻灯片", Toast.ToastType.Warning);
                    return;
                }

                // 获取幻灯片尺寸
                var pageSetup = app.ActivePresentation?.PageSetup;
                float slideWidth = pageSetup?.SlideWidth ?? 0;
                float slideHeight = pageSetup?.SlideHeight ?? 0;

                if (slideWidth <= 0 || slideHeight <= 0)
                {
                    Toast.Show("无法获取幻灯片尺寸", Toast.ToastType.Warning);
                    return;
                }

                List<NETOP.Shape> createdShapes = [];
                string successMessage = "";

                // 1. 处理选中形状
                if (sel != null)
                {
                    // 处理单个形状
                    if (sel is NETOP.Shape shape)
                    {
                        var (top, left, bottom, right) = ShapeUtils.GetShapeBorderWeights(shape);

                        // 计算矩形参数
                        float rectLeft = shape.Left - left;
                        float rectTop = shape.Top - top;
                        float rectWidth = shape.Width + (left + right);
                        float rectHeight = shape.Height + (top + bottom);

                        // 创建矩形
                        var rect = ShapeUtils.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, shape.Rotation);
                        if (rect != null) createdShapes.Add(rect);
                    }
                    // 处理形状范围
                    else if (sel is NETOP.ShapeRange shapes)
                    {
                        if (shapes.Count > 0)
                        {
                            for (int i = 1; i <= shapes.Count; i++)
                            {
                                var currentShape = shapes[i];
                                var (top, left, bottom, right) = ShapeUtils.GetShapeBorderWeights(currentShape);

                                // 计算矩形参数
                                float rectLeft = currentShape.Left - left;
                                float rectTop = currentShape.Top - top;
                                float rectWidth = currentShape.Width + (left + right);
                                float rectHeight = currentShape.Height + (top + bottom);

                                // 创建矩形
                                var rect = ShapeUtils.AddOneShape(currentSlide, rectLeft, rectTop, rectWidth, rectHeight, currentShape.Rotation);

                                if (rect != null) createdShapes.Add(rect);
                            }
                        }
                    }

                    if (createdShapes.Count > 0)
                    {
                        var shapeNames = createdShapes.Select(s => s.Name).ToArray();
                        currentSlide.Shapes.Range(shapeNames).Select();
                        successMessage = $"已为 {createdShapes.Count} 个形状创建外框";
                    }
                }
                // 2. 处理选中幻灯片 和 无选中
                else
                {
                    // 创建要处理的幻灯片列表
                    List<NETOP.Slide> slidesToProcess = [];
                    // 检查是否选中了幻灯片
                    var window = app.ActiveWindow;
                    if (window != null && window.Selection?.Type == NETOP.Enums.PpSelectionType.ppSelectionSlides)
                    {
                        // 选中幻灯片的情况
                        var slideRange = window.Selection.SlideRange;
                        if (slideRange?.Count > 0)
                        {
                            for (int i = 1; i <= slideRange.Count; i++)
                            {
                                slidesToProcess.Add(slideRange[i]);
                            }
                            successMessage = $"已在 {slideRange.Count} 张幻灯片上创建矩形";
                        }
                    }
                    else
                    {
                        // 无选中的情况
                        slidesToProcess.Add(currentSlide);
                        successMessage = "已创建页面大小的矩形";
                    }

                    // 统一处理幻灯片列表
                    if (slidesToProcess.Count > 0)
                    {
                        for (int i = 0; i < slidesToProcess.Count; i++)
                        {
                            var slide = slidesToProcess[i];
                            var rect = ShapeUtils.AddOneShape(slide, 0, 0, slideWidth, slideHeight);

                            if (rect != null)
                            {
                                createdShapes.Add(rect);
                                // 如果是第一张幻灯片，则选中其上的矩形
                                if (i == 0) rect.Select();
                            }
                        }
                    }
                }

                // 显示结果消息
                if (createdShapes.Count > 0)
                {
                    Toast.Show(successMessage, Toast.ToastType.Success);
                }
                else
                {
                    Toast.Show("未创建任何矩形", Toast.ToastType.Info);
                }
            });
        }

        public static void ExecuteAlignment(NETOP.Application app, AlignmentType alignment, bool alignToSlideMode)
        {
            app.StartNewUndoEntry(); // 开始新的撤销单元
            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                if (sel == null)
                {
                    Toast.Show("未选择形状", Toast.ToastType.Warning);
                    return;
                }

                NETOP.ShapeRange shapes;
                // 尝试直接转换为 ShapeRange
                if (sel is NETOP.ShapeRange shapeRange)
                {
                    shapes = shapeRange;
                }
                // 如果不是，则尝试处理单个 Shape
                else if (sel is NETOP.Shape shape && shape.Parent is NETOP.Slide parentSlide)
                {
                    // 使用模式匹配确保 Parent 是 Slide 类型，然后创建 ShapeRange
                    shapes = parentSlide.Shapes.Range(new object[] { shape.Name });
                }
                else
                {
                    // 如果两种情况都不满足，则选择无效，直接返回
                    Toast.Show("选择无效", Toast.ToastType.Warning);
                    return;
                }

                // 判断对齐基准，1.单选形状：总是对齐到幻灯片；2.多选形状：根据按钮状态决定
                MsoTriState alignToSlide = (shapes.Count == 1 || alignToSlideMode) ? MsoTriState.msoTrue : MsoTriState.msoFalse;

                // 创建对齐命令字典
                Dictionary<AlignmentType, Action> alignCommands = new()
                {
                    [AlignmentType.Left] = () => shapes.Align(MsoAlignCmd.msoAlignLefts, alignToSlide),
                    [AlignmentType.Right] = () => shapes.Align(MsoAlignCmd.msoAlignRights, alignToSlide),
                    [AlignmentType.Top] = () => shapes.Align(MsoAlignCmd.msoAlignTops, alignToSlide),
                    [AlignmentType.Bottom] = () => shapes.Align(MsoAlignCmd.msoAlignBottoms, alignToSlide),
                    [AlignmentType.Centers] = () => shapes.Align(MsoAlignCmd.msoAlignCenters, alignToSlide),
                    [AlignmentType.Middles] = () => shapes.Align(MsoAlignCmd.msoAlignMiddles, alignToSlide),
                    [AlignmentType.Horizontally] = () =>
                    {
                        // 根据对齐基准确定所需的最小形状数
                        int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
                        // 检查数量，如果满足则执行，否则抛出明确的异常
                        if (shapes.Count >= minRequired)
                        {
                            shapes.Distribute(MsoDistributeCmd.msoDistributeHorizontally, alignToSlide);
                        }
                        else
                        {
                            string basis = (alignToSlide == MsoTriState.msoTrue) ? "页面" : "形状";
                            throw new InvalidOperationException($"{basis}基准至少选{minRequired}个。");
                        }
                    },
                    [AlignmentType.Vertically] = () =>
                    {
                        // 根据对齐基准确定所需的最小形状数
                        int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
                        // 检查数量，如果满足则执行，否则抛出明确的异常
                        if (shapes.Count >= minRequired)
                        {
                            shapes.Distribute(MsoDistributeCmd.msoDistributeVertically, alignToSlide);
                        }
                        else
                        {
                            string basis = (alignToSlide == MsoTriState.msoTrue) ? "页面" : "形状";
                            throw new InvalidOperationException($"{basis}基准至少选{minRequired}个。");
                        }
                    }
                };

                // 执行对齐操作
                if (alignCommands.TryGetValue(alignment, out var command))
                {
                    try
                    {
                        command();
                        Toast.Show("对齐成功.", Toast.ToastType.Success);
                    }
                    catch (InvalidOperationException ex)
                    {
                        // 显示分布操作所需的形状数量不足的提示
                        Toast.Show(ex.Message, Toast.ToastType.Warning);
                    }
                    catch (Exception ex)
                    {
                        // 捕获其他可能的异常
                        Toast.Show($"对齐失败{ex.Message}", Toast.ToastType.Error);
                    }
                }
                else
                {
                    var errorMsg = $"未知对齐方式: {alignment}";
                    Toast.Show(errorMsg, Toast.ToastType.Error);
                }
            });
        }

        /// <summary>
        /// 隐藏/显示对象：选中对象时隐藏选中对象，无选中对象时显示所有对象。
        /// </summary>
        /// <param name="app">PowerPoint 应用程序实例。</param>
        public static void ToggleShapeVisibility(NETOP.Application app)
        {
            ExHandler.Run(() =>
            {
                var slide = TryGetCurrentSlide(app);
                if (slide == null)
                {
                    Toast.Show("未找到当前幻灯片", Toast.ToastType.Warning);
                    return;
                }

                var sel = ShapeUtils.ValidateSelection(app);
                if (sel != null)
                {
                    // --- 场景1: 隐藏选中的对象 ---
                    if (sel is NETOP.Shape shape)
                    {
                        // 单个形状的情况，创建临时ShapeRange
                        List<NETOP.Shape> shapeList = [shape];
                        app.StartNewUndoEntry();
                        try
                        {
                            shape.Visible = MsoTriState.msoFalse;
                            Toast.Show("已隐藏 1 个对象", Toast.ToastType.Success);
                        }
                        finally
                        {
                            shapeList.DisposeAll();
                        }
                    }
                    else if (sel is NETOP.ShapeRange shapeRange)
                    {
                        // 多个形状的情况
                        HideSelectedShapes(app, shapeRange);
                    }
                }
                else
                {
                    // --- 场景2: 显示所有对象 ---
                    ShowAllHiddenShapes(app, slide.Shapes);
                }
            });
        }

        /// <summary>
        /// 隐藏指定形状范围内的所有形状。
        /// </summary>
        /// <param name="shapeRange">要隐藏的形状范围。</param>
        private static void HideSelectedShapes(NETOP.Application app, NETOP.ShapeRange shapeRange)
        {
            // 使用目标类型 new() 和集合表达式 [] (C# 9.0+ & C# 12.0)
            List<NETOP.Shape> shapesToHide = new(shapeRange.Count);
            for (int i = 1; i <= shapeRange.Count; i++)
            {
                shapesToHide.Add(shapeRange[i]);
            }

            app.StartNewUndoEntry(); // 将撤销操作移到具体动作前，更精确
            try
            {
                foreach (var shape in shapesToHide)
                {
                    shape.Visible = MsoTriState.msoFalse;
                }
                Toast.Show($"已隐藏 {shapesToHide.Count} 个对象", Toast.ToastType.Success);
            }
            finally
            {
                shapesToHide.DisposeAll();
            }
        }

        /// <summary>
        /// 显示幻灯片上所有被隐藏的形状。
        /// </summary>
        /// <param name="shapes">幻灯片的形状集合。</param>
        private static void ShowAllHiddenShapes(NETOP.Application app, NETOP.Shapes shapes)
        {
            List<NETOP.Shape> shapesToShow = [];

            // 1. 找出所有需要显示的对象
            for (int i = 1; i <= shapes.Count; i++)
            {
                var shape = shapes[i];
                if (shape.Visible == MsoTriState.msoFalse)
                {
                    shapesToShow.Add(shape);
                }
            }

            // 2. 根据列表内容执行操作和反馈
            if (shapesToShow.Count > 0)
            {
                app.StartNewUndoEntry();
                try
                {
                    foreach (var shape in shapesToShow)
                    {
                        shape.Visible = MsoTriState.msoTrue;
                    }
                    Toast.Show($"已显示 {shapesToShow.Count} 个对象", Toast.ToastType.Success);
                }
                finally
                {
                    shapesToShow.DisposeAll();
                }
            }
            else
            {
                Toast.Show("没有需要显示的对象", Toast.ToastType.Info);
            }
        }

        #endregion Public Methods
    }
}
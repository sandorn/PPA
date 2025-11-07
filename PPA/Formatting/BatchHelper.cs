using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using PPA.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NETOP = NetOffice.PowerPointApi;
//using NETOP.Enums;

namespace PPA.Formatting
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

        /// <summary>
        /// 同步美化表格（原始版本，保留用于向后兼容）
        /// </summary>
        /// <remarks>
        /// <para>
        /// 注意：此方法为同步版本，在执行耗时操作时会阻塞 PowerPoint UI 线程。
        /// 建议使用异步版本 <see cref="Bt501_ClickAsync"/> 以获得更好的用户体验。
        /// </para>
        /// <para>
        /// 保留此方法的原因：
        /// 1. 向后兼容：确保现有代码调用不受影响
        /// 2. 简单场景：对于少量表格，同步执行可能更简单
        /// 3. 调试方便：同步代码更容易调试和追踪
        /// </para>
        /// </remarks>
        /// <param name="app">PowerPoint 应用程序实例</param>
        public static void Bt501_Click(NETOP.Application app)
        {
            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.FormatTables);
            var slide = TryGetCurrentSlide(app);

            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                int tableCount = 0;

                if (sel != null)
                {
                    // 有选中对象的情况，美化选中对象中的表格
                    if (sel is NETOP.Shape shape)
                    {
                        if (shape.HasTable == MsoTriState.msoTrue)
                        {
                            FormatHelper.FormatTables(shape.Table);
                            tableCount++;
                        }
                    }
                    else if (sel is NETOP.ShapeRange shapes)
                    {
                        foreach (NETOP.Shape s in shapes)
                        {
                            if (s.HasTable == MsoTriState.msoTrue)
                            {
                                FormatHelper.FormatTables(s.Table);
                                tableCount++;
                            }
                        }
                    }

                    if (tableCount > 0)
                        Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success", "成功美化 {0} 个表格", tableCount), Toast.ToastType.Success);
                    else
                        Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoSelection", "选中的对象中没有表格"), Toast.ToastType.Info);
                }
                else
                {
                    // 未选中对象的情况，美化当前幻灯片所有表格
                    if (slide != null)
                    {
                        foreach (NETOP.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTable == MsoTriState.msoTrue)
                            {
                                FormatHelper.FormatTables(shape.Table);
                                tableCount++;
                            }
                        }

                        if (tableCount > 0)
                            Toast.Show(ResourceManager.GetString("Toast_FormatTables_Success", "成功美化 {0} 个表格", tableCount), Toast.ToastType.Success);
                        else
                            Toast.Show(ResourceManager.GetString("Toast_FormatTables_NoTables", "当前幻灯片上没有表格"), Toast.ToastType.Info);
                    }
                }
            }, enableTiming: true);
        }

        /// <summary>
        /// 异步美化表格（支持进度报告和取消）
        /// </summary>
        /// <remarks>
        /// <para>
        /// 这是 <see cref="Bt501_Click"/> 的异步版本，提供以下改进：
        /// 1. 非阻塞执行：不会冻结 PowerPoint UI 线程
        /// 2. 进度反馈：通过 <paramref name="progress"/> 参数报告美化进度
        /// 3. 取消支持：通过 <paramref name="cancellationToken"/> 支持取消操作
        /// 4. 更好的用户体验：执行耗时操作时用户可以继续使用 PowerPoint
        /// </para>
        /// <para>
        /// 使用场景：
        /// - 美化大量表格时（推荐使用此异步版本）
        /// - 需要进度反馈时
        /// - 需要支持取消操作时
        /// </para>
        /// <para>
        /// 注意：所有 Office COM 对象操作都在 UI 线程执行，确保线程安全。
        /// </para>
        /// </remarks>
        /// <param name="app">PowerPoint 应用程序实例</param>
        /// <param name="progress">进度报告对象，用于报告美化进度（可选）</param>
        /// <param name="cancellationToken">取消令牌，用于取消正在进行的操作（可选）</param>
        /// <returns>表示异步操作的 Task</returns>
        public static async Task Bt501_ClickAsync(
            NETOP.Application app,
            IProgress<AsyncProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            try
            {
                // 必须在 UI 线程执行撤销操作
                await UndoHelper.BeginUndoEntryAsync(app, UndoHelper.UndoNames.FormatTables);

                var slide = await AsyncOperationHelper.RunOnUIThread(() =>
                {
                    return TryGetCurrentSlide(app);
                });

                if (slide == null)
                {
                    Toast.Show(ResourceManager.GetString("Toast_NoSlide"), Toast.ToastType.Warning);
                    return;
                }

                // 在 UI 线程收集表格信息
                var tables = await AsyncOperationHelper.RunOnUIThread(() =>
                {
                    var sel = ShapeUtils.ValidateSelection(app);
                    var result = new List<(NETOP.Shape shape, NETOP.Table table)>();

                    if (sel != null)
                    {
                        // 有选中对象的情况，收集选中对象中的表格
                        if (sel is NETOP.Shape shape && shape.HasTable == MsoTriState.msoTrue)
                        {
                            result.Add((shape, shape.Table));
                            progress?.Report(new AsyncProgress(10, ResourceManager.GetString("Progress_TableFound", "发现表格"), 1, 1));
                        }
                        else if (sel is NETOP.ShapeRange shapes)
                        {
                            int count = 0;
                            foreach (NETOP.Shape s in shapes)
                            {
                                if (s.HasTable == MsoTriState.msoTrue)
                                {
                                    result.Add((s, s.Table));
                                    count++;
                                    progress?.Report(new AsyncProgress(
                                        10,
                                        ResourceManager.GetString("Progress_TableFound_Count", "发现表格 {0}", count),
                                        count,
                                        shapes.Count));
                                }
                            }
                        }
                    }
                    else
                    {
                        // 未选中对象的情况，收集当前幻灯片所有表格
                        foreach (NETOP.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTable == MsoTriState.msoTrue)
                            {
                                result.Add((shape, shape.Table));
                            }
                        }
                    }

                    return result;
                });

                cancellationToken.ThrowIfCancellationRequested();

                int total = tables.Count;

                // 如果没有找到表格
                if (total == 0)
                {
                    var hasSelection = await AsyncOperationHelper.RunOnUIThread(() =>
                    {
                        return ShapeUtils.ValidateSelection(app) != null;
                    });
                    Toast.Show(hasSelection ? ResourceManager.GetString("Toast_FormatTables_NoSelection", "选中的对象中没有表格") : ResourceManager.GetString("Toast_FormatTables_NoTables", "当前幻灯片上没有表格"), Toast.ToastType.Info);
                    return;
                }

                // 报告找到的表格数量（简化日志，不逐条扫描）
                progress?.Report(new AsyncProgress(10, ResourceManager.GetString("Progress_TablesFound", "发现 {0} 个表格", total), total, total));

                progress?.Report(new AsyncProgress(20, ResourceManager.GetString("Progress_FormatTables_Start", "开始美化 {0} 个表格", total), 0, total));

                // 在 UI 线程逐个美化表格（同步执行，但允许 UI 更新）
                for (int i = 0; i < total; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    NETOP.Shape shape = tables[i].shape;
                    NETOP.Table table = tables[i].table;

                    // 计算当前表格的进度（20% 开始，80% 结束）
                    int currentProgress = 20 + (int)((i * 60.0) / total);
                    progress?.Report(new AsyncProgress(
                        currentProgress,
                        ResourceManager.GetString("Progress_FormatTable_Progress", "美化表格 {0}/{1}", i + 1, total),
                        i + 1,
                        total));

                    // 在 UI 线程执行美化
                    await AsyncOperationHelper.RunOnUIThread(() =>
                    {
                        FormatHelper.FormatTables(table);
                    });

                    // 允许 UI 更新（每处理一个表格后）
                    await Task.Delay(10, cancellationToken);
                }

                progress?.Report(new AsyncProgress(100, ResourceManager.GetString("Progress_FormatTables_Complete", "成功美化 {0} 个表格", total), total, total));
            }
            catch
            {
                throw; // 重新抛出异常，让上层处理
            }
        }

        public static void Bt502_Click(NETOP.Application app)
        {
            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.FormatText);

            ExHandler.Run(() =>
            {
                // 获取选中的形状
                var selection = ShapeUtils.ValidateSelection(app);

                // 如果没有选中对象，显示提示并返回
                if (selection == null)
                {
                    Toast.Show(ResourceManager.GetString("Toast_FormatText_NoSelection"), Toast.ToastType.Warning);
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

                // 如果成功美化了文本，显示成功提示
                if (hasFormatted)
                {
                    Toast.Show(ResourceManager.GetString("Toast_FormatText_Success"), Toast.ToastType.Success);
                }
                else
                {
                    Toast.Show(ResourceManager.GetString("Toast_FormatText_NoText"), Toast.ToastType.Warning);
                }
            });
        }

        public static void Bt503_Click(NETOP.Application app)
        {
            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.FormatCharts);

            ExHandler.Run(() =>
            {
                var slide = TryGetCurrentSlide(app);
                if (slide == null) return;

                // 收集需要处理的图表形状
                var chartShapes = new List<NETOP.Shape>();
                var sel = ShapeUtils.ValidateSelection(app);

                if (sel != null)
                {
                    // 处理选中的对象
                    if (sel is NETOP.Shape shape && shape.HasChart == MsoTriState.msoTrue)
                    {
                        chartShapes.Add(shape);
                    }
                    else if (sel is NETOP.ShapeRange shapes)
                    {
                        foreach (NETOP.Shape s in shapes)
                        {
                            if (s.HasChart == MsoTriState.msoTrue)
                                chartShapes.Add(s);
                        }
                    }
                }
                else
                {
                    // 处理当前幻灯片上所有对象
                    foreach (NETOP.Shape shape in slide.Shapes)
                    {
                        if (shape.HasChart == MsoTriState.msoTrue)
                            chartShapes.Add(shape);
                    }
                }

                // 格式化所有图表
                foreach (var shape in chartShapes)
                {
                    FormatHelper.FormatChartText(shape);
                }

                // 显示结果
                if (chartShapes.Count > 0)
                    Toast.Show(ResourceManager.GetString("Toast_FormatCharts_Success", chartShapes.Count), Toast.ToastType.Success);
                else
                    Toast.Show(sel != null ? ResourceManager.GetString("Toast_FormatCharts_NoSelection") : ResourceManager.GetString("Toast_FormatCharts_NoCharts"), Toast.ToastType.Info);
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
            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.CreateBoundingBox);

            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                var currentSlide = TryGetCurrentSlide(app);

                if (currentSlide == null)
                {
                    Toast.Show(ResourceManager.GetString("Toast_NoSlide"), Toast.ToastType.Warning);
                    return;
                }

                // 获取幻灯片尺寸
                var pageSetup = app.ActivePresentation?.PageSetup;
                float slideWidth = pageSetup?.SlideWidth ?? 0;
                float slideHeight = pageSetup?.SlideHeight ?? 0;

                if (slideWidth <= 0 || slideHeight <= 0)
                {
                    Toast.Show(ResourceManager.GetString("Toast_NoSlideSize"), Toast.ToastType.Warning);
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
                        successMessage = ResourceManager.GetString("Toast_CreateBoundingBox_Shapes", createdShapes.Count);
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
                            successMessage = ResourceManager.GetString("Toast_CreateBoundingBox_Slides", slideRange.Count);
                        }
                    }
                    else
                    {
                        // 无选中的情况
                        slidesToProcess.Add(currentSlide);
                        successMessage = ResourceManager.GetString("Toast_CreateBoundingBox_PageSize");
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
                    Toast.Show(ResourceManager.GetString("Toast_CreateBoundingBox_None"), Toast.ToastType.Info);
                }
            });
        }

        public static void ExecuteAlignment(NETOP.Application app, AlignmentType alignment, bool alignToSlideMode)
        {
            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.AlignShapes);
            ExHandler.Run(() =>
            {
                var sel = ShapeUtils.ValidateSelection(app);
                if (sel == null)
                {
                    Toast.Show(ResourceManager.GetString("Toast_NoSelection"), Toast.ToastType.Warning);
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
                    Toast.Show(ResourceManager.GetString("Toast_InvalidSelection"), Toast.ToastType.Warning);
                    return;
                }

                // 判断对齐基准，1.单选形状：总是对齐到幻灯片；2.多选形状：根据按钮状态决定
                MsoTriState alignToSlide = (shapes.Count == 1 || alignToSlideMode) ? MsoTriState.msoTrue : MsoTriState.msoFalse;

                // 执行对齐操作
                switch (alignment)
                {
                    case AlignmentType.Left:
                        shapes.Align(MsoAlignCmd.msoAlignLefts, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Right:
                        shapes.Align(MsoAlignCmd.msoAlignRights, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Top:
                        shapes.Align(MsoAlignCmd.msoAlignTops, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Bottom:
                        shapes.Align(MsoAlignCmd.msoAlignBottoms, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Centers:
                        shapes.Align(MsoAlignCmd.msoAlignCenters, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Middles:
                        shapes.Align(MsoAlignCmd.msoAlignMiddles, alignToSlide);
                        Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                        break;

                    case AlignmentType.Horizontally:
                        {
                            // 根据对齐基准确定所需的最小形状数
                            int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
                            if (shapes.Count >= minRequired)
                            {
                                shapes.Distribute(MsoDistributeCmd.msoDistributeHorizontally, alignToSlide);
                                Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                            }
                            else
                            {
                                string basis = (alignToSlide == MsoTriState.msoTrue) 
                                    ? ResourceManager.GetString("Toast_Basis_Page", "页面") 
                                    : ResourceManager.GetString("Toast_Basis_Shape", "形状");
                                Toast.Show(ResourceManager.GetString("Toast_AlignMinShapes", basis, minRequired), Toast.ToastType.Warning);
                            }
                        }
                        break;

                    case AlignmentType.Vertically:
                        {
                            // 根据对齐基准确定所需的最小形状数
                            int minRequired = (alignToSlide == MsoTriState.msoTrue) ? 1 : 3;
                            if (shapes.Count >= minRequired)
                            {
                                shapes.Distribute(MsoDistributeCmd.msoDistributeVertically, alignToSlide);
                                Toast.Show(ResourceManager.GetString("Toast_AlignSuccess"), Toast.ToastType.Success);
                            }
                            else
                            {
                                string basis = (alignToSlide == MsoTriState.msoTrue) 
                                    ? ResourceManager.GetString("Toast_Basis_Page", "页面") 
                                    : ResourceManager.GetString("Toast_Basis_Shape", "形状");
                                Toast.Show(ResourceManager.GetString("Toast_AlignMinShapes", basis, minRequired), Toast.ToastType.Warning);
                            }
                        }
                        break;

                    default:
                        Toast.Show(ResourceManager.GetString("Toast_UnknownAlignment", alignment.ToString()), Toast.ToastType.Error);
                        break;
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
                    Toast.Show(ResourceManager.GetString("Toast_NoSlide"), Toast.ToastType.Warning);
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
                        UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.HideShapes);
                        try
                        {
                            shape.Visible = MsoTriState.msoFalse;
                            Toast.Show(ResourceManager.GetString("Toast_HideShapes_Single"), Toast.ToastType.Success);
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

            UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.HideShapes);
            try
            {
                foreach (var shape in shapesToHide)
                {
                    shape.Visible = MsoTriState.msoFalse;
                }
                Toast.Show(ResourceManager.GetString("Toast_HideShapes_Multiple", shapesToHide.Count), Toast.ToastType.Success);
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
                UndoHelper.BeginUndoEntry(app, UndoHelper.UndoNames.ShowShapes);
                try
                {
                    foreach (var shape in shapesToShow)
                    {
                        shape.Visible = MsoTriState.msoTrue;
                    }
                    Toast.Show(ResourceManager.GetString("Toast_ShowShapes_Multiple", shapesToShow.Count), Toast.ToastType.Success);
                }
                finally
                {
                    shapesToShow.DisposeAll();
                }
            }
            else
            {
                Toast.Show(ResourceManager.GetString("Toast_ShowShapes_None"), Toast.ToastType.Info);
            }
        }

        #endregion Public Methods
    }
}
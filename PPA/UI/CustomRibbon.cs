using PPA.Helpers;
using PPA.MSOAPI;
using PPA.Properties;
using Project.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using ALT = PPA.Helpers.BatchHelper.AlignmentType;
using NETOP = NetOffice.PowerPointApi;
using Office = Microsoft.Office.Core;

namespace PPA
{
    [ComVisible(true)]
    public class CustomRibbon : Office.IRibbonExtensibility, IDisposable
    {
        #region Private Fields

        private Office.IRibbonUI _ribbonUI;
        private NETOP.Application _app;
        private readonly Dictionary<string, Bitmap> _iconCache;
        private bool _tb101Press;
        private bool _disposed = false;
        private bool _appInitialized = false;

        #endregion Private Fields

        #region Initialization & Setup

        /// <summary>
        /// CustomRibbon 类的构造函数
        /// </summary>
        public CustomRibbon()
        {
            Profiler.LogMessage("[Ribbon] 构造...");
            _iconCache = [];
            _tb101Press = false;
            // 注意：此时不初始化 _app，等待 SetApplication 调用
        }

        /// <summary>
        /// 在 ThisAddIn Startup 完成后设置 Application 对象
        /// </summary>
        /// <param name="application">PowerPoint Application 实例</param>
        public void SetApplication(NETOP.Application application)
        {
            if (application == null)
            {
                Profiler.LogMessage("[Ribbon] SetApplication 传入空 Application 对象");
                return;
            }

            _app = application;
            _appInitialized = true;
            Profiler.LogMessage("[Ribbon] Application 设置成功");
        }

        /// <summary>
        /// Ribbon UI 加载时调用的事件处理器
        /// </summary>
        /// <param name="ribbonUI">功能区UI接口</param>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                // Profiler.LogMessage(UI初始化...");
                _ribbonUI = ribbonUI;
                PreloadIcons();
                _tb101Press = false;

                _ribbonUI?.Invalidate();
                Profiler.LogMessage("UI加载成功");
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"UI加载错误: {ex.Message}");
            }
        }

        /// <summary>
        /// IRibbonExtensibility 接口的实现，用于加载 Ribbon XML
        /// </summary>
        /// <param name="ribbonID">功能区标识符</param>
        /// <returns>Ribbon的XML字符串</returns>
        public string GetCustomUI(string ribbonID)
        {
            // Profiler.LogMessage($"ribbonID: {ribbonID}");

            try
            {
                string ribbonXml = LoadRibbonXmlFromFile();
                if (!string.IsNullOrEmpty(ribbonXml))
                {
                    return ribbonXml;
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"从文件加载XML失败: {ex.Message}");
            }

            return Resources.RibbonXml;
        }

        /// <summary>
        /// 预加载所有 Ribbon 图标到缓存中，提高UI响应速度
        /// </summary>
        private void PreloadIcons()
        {
            if (_iconCache.Count > 0) return;

            // Profiler.LogMessage("预加载图标...");

            try
            {
                Dictionary<string, Bitmap> icons = new()
                {
                    ["Tb101_1"] = Properties.Resources.slide,
                    ["Tb101_0"] = Properties.Resources.shap,
                    ["Bt121"] = Properties.Resources.Bt121,
                    ["Bt122"] = Properties.Resources.Bt122,
                    ["Bt123"] = Properties.Resources.Bt123,
                    ["Bt124"] = Properties.Resources.Bt124,
                    ["Bt204"] = Properties.Resources.Bt204,
                    ["Bt211"] = Properties.Resources.Bt211,
                    ["Bt212"] = Properties.Resources.Bt212,
                    ["Bt213"] = Properties.Resources.Bt213,
                    ["Bt214"] = Properties.Resources.Bt214,
                    ["Bt301"] = Properties.Resources.Bt301,
                    ["Bt302"] = Properties.Resources.Bt302,
                    ["Bt303"] = Properties.Resources.Bt303,
                    ["Bt311"] = Properties.Resources.Bt311,
                    ["Bt312"] = Properties.Resources.Bt312,
                    ["Bt313"] = Properties.Resources.Bt313,
                    ["Bt321"] = Properties.Resources.Bt321,
                    ["Bt322"] = Properties.Resources.Bt323,
                    ["Bt323"] = Properties.Resources.Bt322,
                    ["Bt401"] = Properties.Resources.Bt401,
                    ["Bt402"] = Properties.Resources.Bt402,
                    ["Bt601"] = Properties.Resources.Bt601
                };

                foreach (var icon in icons)
                {
                    _iconCache[icon.Key] = icon.Value;
                }

                // Profiler.LogMessage($"已预加载 {_iconCache.Count} 个图标");
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"预加载图标错误: {ex.Message}");
            }
        }

        #endregion Initialization & Setup


        #region State & Property Getters

        /// <summary>
        /// 获取 Ribbon 控件的图标
        /// </summary>
        public Bitmap GetIcon(Office.IRibbonControl control)
        {
            try
            {
                string itemId = control.Id;
                if (control.Id == "Tb101")
                {
                    itemId = _tb101Press ? "Tb101_1" : "Tb101_0";
                }

                if (_iconCache.TryGetValue(itemId, out Bitmap bmp))
                {
                    return bmp;
                }

                Profiler.LogMessage($"未找到图标: {itemId}");
                return null;
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"获取图标错误 | {control.Id}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取切换按钮的标签
        /// </summary>
        /// <param name="control">功能区控件对象</param>
        /// <returns>切换按钮的显示文本</returns>
        public string GetTbLabel(Office.IRibbonControl control)
        {
            // Profiler.LogMessage($"获取切换按钮标签 | {control.Id}");

            return control.Id switch
            {
                "Tb101" => _tb101Press ? "页面" : "形状",
                _ => string.Empty,
            };
        }

        /// <summary>
        /// 获取切换按钮的按下状态
        /// </summary>
        /// <param name="control">功能区控件对象</param>
        /// <returns>切换按钮的当前状态</returns>
        public bool GetTbState(Office.IRibbonControl control)
        {
            Profiler.LogMessage($"获取切换按钮状态: {control.Id}");

            return control.Id switch
            {
                "Tb101" => _tb101Press,
                _ => false,
            };
        }

        #endregion State & Property Getters

        #region Event Handlers

        /// <summary>
        /// 处理普通按钮的点击事件
        /// </summary>
        /// <param name="control">功能区控件对象</param>
        public void OnAction(Office.IRibbonControl control)
        {
            Profiler.LogMessage($"按钮点击事件: {control.Id}");

            if (!_appInitialized || _app == null)
            {
                Profiler.LogMessage($"Application 未初始化，跳过操作: {control.Id}");
                return;
            }

            try
            {
                ExecuteButtonAction(control.Id);
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"按钮操作错误 {control.Id}: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理切换按钮的点击事件
        /// </summary>
        public void TbOnAction(Office.IRibbonControl control, bool pressed)
        {
            // Profiler.LogMessage($"切换按钮点击事件 | {control.Id}, pressed: {pressed}");

            try
            {
                if (control.Id == "Tb101")
                {
                    _tb101Press = pressed;
                    _ribbonUI?.InvalidateControl("Tb101");
                    Profiler.LogMessage($"切换状态 | {control.Id}: {(_tb101Press ? "页面" : "形状")}");
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"切换按钮点击事件错误 | {control.Id}: {ex.Message}");
            }
        }

        #endregion Event Handlers


        #region Core Business Logic

        /// <summary>
        /// 根据按钮 ID 执行相应的业务逻辑
        /// </summary>
        /// <param name="buttonId">按钮标识符</param>
        private void ExecuteButtonAction(string buttonId)
        {
            if (!_appInitialized || _app == null)
            {
                Profiler.LogMessage($"Application 不可用，无法执行操作");
                return;
            }

            switch (buttonId)
            {
                case "Bt101":
                    BatchHelper.ExecuteAlignment(_app, ALT.Left, _tb101Press);
                    break;
                case "Bt102":
                    BatchHelper.ExecuteAlignment(_app, ALT.Centers, _tb101Press);
                    break;
                case "Bt103":
                    BatchHelper.ExecuteAlignment(_app, ALT.Right, _tb101Press);
                    break;
                case "Bt104":
                    BatchHelper.ExecuteAlignment(_app, ALT.Horizontally, _tb101Press);
                    break;
                case "Bt111":
                    BatchHelper.ExecuteAlignment(_app, ALT.Top, _tb101Press);
                    break;
                case "Bt112":
                    BatchHelper.ExecuteAlignment(_app, ALT.Middles, _tb101Press);
                    break;
                case "Bt113":
                    BatchHelper.ExecuteAlignment(_app, ALT.Bottom, _tb101Press);
                    break;
                case "Bt114":
                    BatchHelper.ExecuteAlignment(_app, ALT.Vertically, _tb101Press);
                    break;
                case "Bt121":
                    AlignHelper.AttachLeft(_app);
                    break;
                case "Bt122":
                    AlignHelper.AttachRight(_app);
                    break;
                case "Bt123":
                    AlignHelper.AttachTop(_app);
                    break;
                case "Bt124":
                    AlignHelper.AttachBottom(_app);
                    break;
                case "Bt201":
                    AlignHelper.SetEqualWidth(_app);
                    break;
                case "Bt202":
                    AlignHelper.SetEqualHeight(_app);
                    break;
                case "Bt203":
                    AlignHelper.SetEqualSize(_app);
                    break;
                case "Bt204":
                    AlignHelper.SwapSize(_app);
                    break;
                case "Bt211":
                    AlignHelper.StretchLeft(_app);
                    break;
                case "Bt212":
                    AlignHelper.StretchRight(_app);
                    break;
                case "Bt213":
                    AlignHelper.StretchTop(_app);
                    break;
                case "Bt214":
                    AlignHelper.StretchBottom(_app);
                    break;
                case "Bt301":
                    AlignHelper.GuideAlignLeft(_app);
                    break;
                case "Bt302":
                    AlignHelper.GuideAlignHCenter(_app);
                    break;
                case "Bt303":
                    AlignHelper.GuideAlignRight(_app);
                    break;
                case "Bt311":
                    AlignHelper.GuideAlignTop(_app);
                    break;
                case "Bt312":
                    AlignHelper.GuideAlignVCenter(_app);
                    break;
                case "Bt313":
                    AlignHelper.GuideAlignBottom(_app);
                    break;
                case "Bt321":
                    AlignHelper.GuidesStretchWidth(_app);
                    break;
                case "Bt322":
                    AlignHelper.GuidesStretchHeight(_app);
                    break;
                case "Bt323":
                    AlignHelper.GuidesStretchSize(_app);
                    break;
                case "Bt401":
                    BatchHelper.ToggleShapeVisibility(_app);
                    break;
                case "Bt402":
                    MSOICrop.CropShapesToSlide(_app);
                    break;
                case "Bt501":
                    BatchHelper.Bt501_Click(_app);
                    break;
                case "Bt502":
                    BatchHelper.Bt502_Click(_app);
                    break;
                case "Bt503":
                    BatchHelper.Bt503_Click(_app);
                    break;
                case "Bt601":
                    BatchHelper.Bt601_Click(_app);
                    break;
                default:
                    Profiler.LogMessage($"未知按钮ID: {buttonId}");
                    break;
            }
        }

        #endregion Core Business Logic


        #region Lifecycle Management (IDisposable)

        /// <summary>
        /// 公共的 Dispose 方法
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 受保护的 Dispose 方法，用于释放资源
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                Profiler.LogMessage($"释放资源...");

                foreach (var kvp in _iconCache)
                {
                    try
                    {
                        kvp.Value?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Profiler.LogMessage($"释放图标资源时出错 | {kvp.Key}: {ex.Message}");
                    }
                }
                _iconCache.Clear();

                try
                {
                    if (_ribbonUI != null)
                    {
                        Marshal.ReleaseComObject(_ribbonUI);
                        _ribbonUI = null;
                    }
                }
                catch (Exception ex)
                {
                    Profiler.LogMessage($"释放UI时出错: {ex.Message}");
                }

                // 注意：不释放 _app，因为它由 ThisAddIn 管理
            }

            _disposed = true;
        }

        #endregion Lifecycle Management (IDisposable)


        #region Private Helper Methods

        /// <summary>
        /// 从文件中加载 Ribbon XML，如果找不到则返回 null
        /// 支持从多个可能的路径加载，增强灵活性
        /// </summary>
        /// <returns>加载的XML字符串，如未找到则返回null</returns>
        private string LoadRibbonXmlFromFile()
        {
            string filePath = FileLocator.FindFile("UI\\Ribbon.xml");
            if (filePath != null)
            {
                return File.ReadAllText(filePath);
            }
            Profiler.LogMessage($"未找到外部XML文件，使用嵌入式资源");
            return null;
        }

        #endregion Private Helper Methods
    }
}

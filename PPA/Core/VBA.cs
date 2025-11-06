using NetOffice.VBIDEApi.Enums;
using Project.Utilities;
using System;
using System.Diagnostics;
using System.IO;
using NETOP = NetOffice.PowerPointApi;

namespace VBAApi
{
    public static class VbaExecutor
    {
        #region Public Methods

        public static void ExecuteVbaCode(NETOP.Application app, string vbaCode, string macroName, params object[] args)
        {
            dynamic presentation = app.ActivePresentation;

            ExHandler.Run(() =>
            {
                var vbProject = presentation.VBProject;
                var module = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                var codeModule = module.CodeModule;
                codeModule.AddFromString(vbaCode);
                presentation.Save();    //触发编译
                Profiler.LogMessage($"模块创建：{module.Name}({codeModule.CountOfLines} lines) | args.Length  : {args.Length} ");

                /*—运行宏—*/
                if (args.Length == 0) app.Run($"{module.Name}.{macroName}");
                else app.Run($"{module.Name}.{macroName}", args);

                /*—清理—*/
                //短暂延迟防锁COM对象
                System.Threading.Thread.Sleep(50);
                vbProject.VBComponents.Remove(module);
            }, "ExecuteVbaCode");
        }

        #endregion Public Methods
    }

    public static class VbaManager
    {
        private static bool _isInitialized = false;
        private const string ModuleName = "PPA_CoreFormatter";
        private static string _vbaCodeCache = null; // 添加一个静态缓存

        public static void EnsureModuleInitialized(NETOP.Application app)
        {
            if (_isInitialized) return;

            ExHandler.Run(() =>
            {
                // 从文件读取代码
                string vbaCode = GetVbaCodeFromFile();

                dynamic presentation = app.ActivePresentation;
                var vbProject = presentation.VBProject;
                bool moduleExists = false;

                foreach (dynamic component in vbProject.VBComponents)
                {
                    if (component.Name == ModuleName)
                    {
                        moduleExists = true;
                        break;
                    }
                }

                if (!moduleExists)
                {
                    var module = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                    module.Name = ModuleName;
                    var codeModule = module.CodeModule;
                    codeModule.AddFromString(vbaCode);
                    //presentation.Save();
                    Profiler.LogMessage($"VBA 模块 '{ModuleName}' 已在内存中创建。");
                }
                else
                {
                    Profiler.LogMessage($"VBA 模块 '{ModuleName}' 已存在，跳过创建。");
                }

                _isInitialized = true;

            }, "初始化VBA模块");
        }

        public static void RunMacro(NETOP.Application app, string macroName, params object[] args)
        {
            EnsureModuleInitialized(app); // 确保模块存在

            ExHandler.Run(() =>
            {
                // 直接运行宏，不再创建和删除
                if (args.Length == 0)
                    app.Run($"{ModuleName}.{macroName}");
                else
                    app.Run($"{ModuleName}.{macroName}", args);
            }, "运行VBA宏");
        }

        // 新的文件读取方法
        private static string GetVbaCodeFromFile()
        {
            // 如果已经读取过，直接从缓存返回
            if (_vbaCodeCache != null)
            {
                return _vbaCodeCache;
            }

            try
            {
                // 使用 FileLocator 查找文件
                string vbaFilePath = FileLocator.FindFile("Resources\\TableFormatter.vba");

                if (vbaFilePath != null)
                {
                    _vbaCodeCache = File.ReadAllText(vbaFilePath);
                    //Profiler.LogMessage($"成功从文件加载 VBA 代码: {vbaFilePath}");
                    return _vbaCodeCache;
                }
                else
                {
                    // 如果文件不存在，抛出异常
                    throw new FileNotFoundException($"VBA 代码文件未找到，请确保 'TableFormatter.vba' 已复制到输出目录。");
                }
            }
            catch (Exception ex)
            {
                Profiler.LogMessage($"读取 VBA 文件时出错: {ex.Message}");
                throw; // 重新抛出异常，让 ExHandler 捕获并记录
            }
        }
    }
}
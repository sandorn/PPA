using PPA.Core;
using System;
using System.IO;
using System.Reflection;

namespace PPA.Utilities
{
    public static class FileLocator
    {
        /// <summary>
        /// 在多个可能的位置搜索文件
        /// 搜索优先级为常见的可执行文件位置
        /// </summary>
        /// <param name="relativePath">相对于常见位置的相对路径，如 "Properties\Ribbon.xml" 或 "TableFormatter.vba"</param>
        /// <returns>找到的文件的完整路径，如果未找到则返回 null。</returns>
        public static string FindFile(string relativePath)
        {
            if (string.IsNullOrEmpty(relativePath))
            {
                return null;
            }

            string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                     ?? AppDomain.CurrentDomain.BaseDirectory;

            string[] candidates =
    [
        Path.Combine(baseDir, relativePath),
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath),
        Path.Combine(baseDir, "..", relativePath),
        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", relativePath)
    ];

            for (int i = 0; i < candidates.Length; i++)
            {
                string candidate = candidates[i];
                string fullPath = Path.GetFullPath(candidate);

                if (File.Exists(fullPath))
                {
                    Profiler.LogMessage($"找到文件: {fullPath}");
                    return fullPath;
                }
            }

            Profiler.LogMessage($"未找到文件: {relativePath}");
            return null;
        }
    }
}

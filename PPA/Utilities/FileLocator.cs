using PPA.Core;
using System;
using System.IO;
using System.Reflection;

namespace PPA.Utilities
{
	public static class FileLocator
	{
		/// <summary>
		/// 在多个可能的位置搜索文件 搜索优先级为常见的可执行文件位置 支持 ClickOnce 部署环境，会自动查找 .deploy 扩展名的文件 注意：Ribbon.xml
		/// 已改为嵌入式资源，不再使用此方法加载 此工具类保留用于将来可能需要从文件系统加载的其他文件
		/// </summary>
		/// <param name="relativePath"> 相对于常见位置的相对路径，如 "UI\Ribbon.xml" 或 "TableFormatter.vba" </param>
		/// <returns> 找到的文件的完整路径，如果未找到则返回 null。 </returns>
		public static string FindFile(string relativePath)
		{
			if(string.IsNullOrEmpty(relativePath))
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

			// 首先尝试查找原始文件名
			for(int i = 0;i<candidates.Length;i++)
			{
				string candidate = candidates[i];
				string fullPath = Path.GetFullPath(candidate);

				if(File.Exists(fullPath))
				{
					Profiler.LogMessage($"找到文件: {fullPath}");
					return fullPath;
				}
			}

			// 如果在 ClickOnce 部署环境中，尝试查找 .deploy 文件 ClickOnce 会将文件重命名为 .deploy 扩展名
			string deployPath = relativePath + ".deploy";
			string[] deployCandidates =
			[
				Path.Combine(baseDir, deployPath),
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, deployPath),
				Path.Combine(baseDir, "..", deployPath),
				Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", deployPath)
			];

			for(int i = 0;i<deployCandidates.Length;i++)
			{
				string candidate = deployCandidates[i];
				string fullPath = Path.GetFullPath(candidate);

				if(File.Exists(fullPath))
				{
					Profiler.LogMessage($"找到 ClickOnce 部署文件: {fullPath}");
					return fullPath;
				}
			}

			Profiler.LogMessage($"未找到文件: {relativePath} (包括 .deploy 版本)");
			return null;
		}
	}
}

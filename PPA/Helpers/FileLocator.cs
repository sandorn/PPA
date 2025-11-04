using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace Project.Utilities // 或者你自己的工具命名空间
{
	public static class FileLocator
	{
		/// <summary>
		/// 查找与插件输出目录相关的文件。
		/// 按优先级顺序在多个可能的位置进行搜索。
		/// </summary>
		/// <param name="relativePath">相对于输出目录的相对路径，例如 "Properties\Ribbon.xml" 或 "TableFormatter.vba"</param>
		/// <returns>找到的文件的完整路径，如果未找到则返回 null。</returns>
		public static string FindFile(string relativePath)
		{
			if(string.IsNullOrEmpty(relativePath))
			{
				return null;
			}

			string baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
					 ?? AppDomain.CurrentDomain.BaseDirectory;

			//Debug.WriteLine($"[FileLocator] 基准目录: {baseDir}");
			//Debug.WriteLine($"[FileLocator] 应用基础目录: {AppDomain.CurrentDomain.BaseDirectory}");

			string[] candidates =
	{
		Path.Combine(baseDir, relativePath),
		Path.Combine(AppDomain.CurrentDomain.BaseDirectory, relativePath),
		Path.Combine(baseDir, "..", relativePath),
		Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", relativePath)
	};

			for(int i = 0;i<candidates.Length;i++)
			{
				string candidate = candidates[i];
				string fullPath = Path.GetFullPath(candidate);
				//Debug.WriteLine($"[FileLocator] 候选路径 {i+1}: {fullPath}");

				if(File.Exists(fullPath))
				{
					Debug.WriteLine($"[FileLocator] 找到文件: {fullPath}");
					return fullPath;
				} else
				{
					//Debug.WriteLine($"[FileLocator] 候选路径 {i+1} 不存在。");
				}
			}

			Debug.WriteLine($"[FileLocator] 所有候选路径均未找到文件: {relativePath}");
			return null;
		}
	}
}

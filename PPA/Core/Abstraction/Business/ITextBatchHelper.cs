using System;
using System.Threading.Tasks;
using PPA.Utilities;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 文本批量操作辅助接口
	/// 注意：当前方法签名待定，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface ITextBatchHelper
	{
		/// <summary>
		/// 批量格式化文本
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		void FormatText(NETOP.Application app);

		/// <summary>
		/// 异步格式化文本（预留，当前实现仍为同步）
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <param name="progress">进度报告回调</param>
		/// <returns>异步任务</returns>
		Task FormatTextAsync(NETOP.Application app, IProgress<AsyncProgress> progress = null);
	}
}


using System;
using System.Threading;
using System.Threading.Tasks;
using PPA.Utilities;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 表格批量操作辅助接口
	/// 注意：当前方法签名待定，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface ITableBatchHelper
	{
		/// <summary>
		/// 同步美化表格
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		void FormatTables(NETOP.Application app);

		/// <summary>
		/// 异步美化表格
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <param name="progress">进度报告回调</param>
		/// <param name="cancellationToken">取消令牌</param>
		/// <returns>表示异步操作的任务</returns>
		Task FormatTablesAsync(
			NETOP.Application app,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default);
	}
}


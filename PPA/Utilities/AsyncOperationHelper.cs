using PPA.Core;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace PPA.Utilities
{
	/// <summary>
	/// 异步操作辅助类 提供统一的异步操作执行框架
	/// </summary>
	public static class AsyncOperationHelper
	{
		private static readonly ILogger _logger = LoggerProvider.GetLogger();

		/// <summary>
		/// 在 UI 线程执行操作
		/// </summary>
		/// <remarks>
		/// 在 Office 插件环境中，Ribbon 事件处理已经在 UI 线程中执行，可以直接同步执行。
		/// 此方法主要作为语义标记，并为了保持接口一致性返回 Task。
		/// </remarks>
		/// <param name="action"> 要在 UI 线程执行的操作 </param>
		public static Task RunOnUIThread(Action action)
		{
			if (action == null)
				throw new ArgumentNullException(nameof(action));

			try
			{
				action();
				return Task.CompletedTask;
			}
			catch (Exception ex)
			{
				return Task.FromException(ex);
			}
		}

		/// <summary>
		/// 执行异步操作，自动处理 UI 线程同步和异常
		/// </summary>
		public static async Task ExecuteAsync(
			Func<IProgress<AsyncProgress>, CancellationToken, Task> operation,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default,
			string operationName = "操作")
		{
			if (operation == null)
				throw new ArgumentNullException(nameof(operation));

			try
			{
				progress?.Report(new AsyncProgress(0, $"开始{operationName}..."));

				await operation(progress, cancellationToken);

				progress?.Report(new AsyncProgress(100, $"{operationName}完成"));
				Toast.Show($"{operationName}完成", Toast.ToastType.Success);
			}
			catch (OperationCanceledException)
			{
				progress?.Report(new AsyncProgress(0, $"{operationName}已取消"));
				Toast.Show($"{operationName}已取消", Toast.ToastType.Info);
			}
			catch (Exception ex)
			{
				ExHandler.Run(() => throw ex, $"{operationName}执行失败");
				Toast.Show($"{operationName}失败: {ex.Message}", Toast.ToastType.Error);
			}
		}

		/// <summary>
		/// 执行异步操作（带返回值）
		/// </summary>
		public static async Task<T> ExecuteAsync<T>(
			Func<IProgress<AsyncProgress>, CancellationToken, Task<T>> operation,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default,
			string operationName = "操作")
		{
			if (operation == null)
				throw new ArgumentNullException(nameof(operation));

			try
			{
				progress?.Report(new AsyncProgress(0, $"开始{operationName}..."));
				var result = await operation(progress, cancellationToken);
				progress?.Report(new AsyncProgress(100, $"{operationName}完成"));
				Toast.Show($"{operationName}完成", Toast.ToastType.Success);
				return result;
			}
			catch (OperationCanceledException)
			{
				progress?.Report(new AsyncProgress(0, $"{operationName}已取消"));
				Toast.Show($"{operationName}已取消", Toast.ToastType.Info);
				throw;
			}
			catch (Exception ex)
			{
				ExHandler.Run(() => throw ex, $"{operationName}执行失败");
				Toast.Show($"{operationName}失败: {ex.Message}", Toast.ToastType.Error);
				throw;
			}
		}

		/// <summary>
		/// 执行异步操作，提供统一的异常处理和进度报告（Fire-and-forget 模式）
		/// </summary>
		public static async void ExecuteAsyncOperation(
			Func<Task> operation,
			string operationName = "异步操作")
		{
			if (operation == null)
				throw new ArgumentNullException(nameof(operation));

			var sw = Stopwatch.StartNew();
			var opName = string.IsNullOrWhiteSpace(operationName) ? "异步操作" : operationName;

			try
			{
				await operation();
				sw.Stop();
				_logger.LogInformation(
					message: $"执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms",
					callerMethod: opName,
					callerFile: string.Empty);
			}
			catch (OperationCanceledException)
			{
				sw.Stop();
				_logger.LogInformation($"{opName}已取消 ({sw.Elapsed.TotalMilliseconds:F0}ms)");
			}
			catch (Exception ex)
			{
				sw.Stop();
				_logger.LogError($"{opName}失败 ({sw.Elapsed.TotalMilliseconds:F0}ms): {ex.GetType().Name}");
				ExHandler.Run(() => throw ex, $"{opName}执行失败");
			}
		}
	}

	/// <summary>
	/// 异步操作进度报告
	/// </summary>
	public class AsyncProgress(int percentage, string message, int currentItem = 0, int totalItems = 0)
	{
		public int Percentage { get; } = Math.Max(0, Math.Min(100, percentage));
		public string Message { get; } = message ?? string.Empty;
		public int CurrentItem { get; } = currentItem;
		public int TotalItems { get; } = totalItems;

		public override string ToString()
		{
			if (TotalItems > 0)
				return $"{Message} ({CurrentItem}/{TotalItems})";
			return $"{Message} ({Percentage}%)";
		}
	}

	/// <summary>
	/// 进度指示器 - 使用 Toast 显示进度
	/// </summary>
	public class ProgressIndicator(string operationName) : IProgress<AsyncProgress>
	{
		private readonly string _operationName = operationName ?? "操作";
		private int _lastPercentage = -1;
		private readonly object _lockObject = new();

		public void Report(AsyncProgress value)
		{
			lock (_lockObject)
			{
				bool shouldUpdate = false;

				if (value.TotalItems > 0)
				{
					shouldUpdate = true;
				}
				else if (value.Percentage / 10 != _lastPercentage / 10)
				{
					shouldUpdate = true;
				}

				if (shouldUpdate)
				{
					_lastPercentage = value.Percentage;

					if (value.Percentage == 0 || value.Percentage == 100 || value.TotalItems > 0)
					{
						string message = value.TotalItems > 0
								? $"{_operationName}: {value.CurrentItem}/{value.TotalItems}"
								: $"{_operationName}: {value.Percentage}%";

						Toast.Show(message, Toast.ToastType.Info, duration: 1000);
					}
				}
			}
		}
	}
}

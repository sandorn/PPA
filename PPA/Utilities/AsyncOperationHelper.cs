using PPA.Core;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPA.Utilities
{
	/// <summary>
	/// 异步操作辅助类 提供统一的异步操作执行框架，确保 COM 对象在 UI 线程操作
	/// </summary>
	public static class AsyncOperationHelper
	{
		/// <summary>
		/// 获取 UI 线程的 SynchronizationContext
		/// </summary>
		private static SynchronizationContext GetUIContext()
		{
			var context = SynchronizationContext.Current;
			if(context==null)
			{
				// Office 插件环境中，可能需要手动设置 SynchronizationContext 尝试使用 Application 的同步上下文，如果不存在则创建新的
				try
				{
					// 在 Office 插件中，通常需要设置 WindowsFormsSynchronizationContext
					if(Application.MessageLoop)
					{
						context=WindowsFormsSynchronizationContext.Current;
					}

					if(context==null)
					{
						context=new WindowsFormsSynchronizationContext();
						SynchronizationContext.SetSynchronizationContext(context);
						Profiler.LogMessage("[AsyncOperationHelper] 创建新的 SynchronizationContext");
					}
				} catch(Exception ex)
				{
					Profiler.LogMessage($"[AsyncOperationHelper] 获取 SynchronizationContext 失败: {ex.Message}");
					// 如果创建失败，使用默认的上下文
					context=new WindowsFormsSynchronizationContext();
				}
			}
			return context;
		}

		/// <summary>
		/// 在 UI 线程执行 COM 操作
		/// </summary>
		/// <remarks>
		/// 在 Office 插件环境中，Ribbon 事件处理已经在 UI 线程中执行， 所以可以直接执行 COM 操作。使用 Task.FromResult 立即返回已完成的任务。
		/// </remarks>
		/// <param name="action"> 要在 UI 线程执行的操作 </param>
		public static Task RunOnUIThread(Action action)
		{
			if(action==null)
				throw new ArgumentNullException(nameof(action));

			try
			{
				// 在 Office 插件中，Ribbon 事件已经在 UI 线程，直接同步执行
				action();
				return Task.CompletedTask;
			} catch(Exception ex)
			{
				return Task.FromException(ex);
			}
		}

		/// <summary>
		/// 在 UI 线程执行 COM 操作（带返回值）
		/// </summary>
		/// <remarks>
		/// 在 Office 插件环境中，Ribbon 事件处理已经在 UI 线程中执行， 所以可以直接执行 COM 操作。使用 Task.FromResult 立即返回已完成的任务。
		/// </remarks>
		/// <typeparam name="T"> 返回值类型 </typeparam>
		/// <param name="func"> 要在 UI 线程执行的函数 </param>
		public static Task<T> RunOnUIThread<T>(Func<T> func)
		{
			if(func==null)
				throw new ArgumentNullException(nameof(func));

			try
			{
				// 在 Office 插件中，Ribbon 事件已经在 UI 线程，直接同步执行
				var result = func();
				return Task.FromResult(result);
			} catch(Exception ex)
			{
				return Task.FromException<T>(ex);
			}
		}

		/// <summary>
		/// 执行异步操作，自动处理 UI 线程同步和异常
		/// </summary>
		/// <param name="operation"> 要执行的操作 </param>
		/// <param name="progress"> 进度报告对象 </param>
		/// <param name="cancellationToken"> 取消令牌 </param>
		/// <param name="operationName"> 操作名称（用于日志和提示） </param>
		public static async Task ExecuteAsync(
			Func<IProgress<AsyncProgress>,CancellationToken,Task> operation,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default,
			string operationName = "操作")
		{
			if(operation==null)
				throw new ArgumentNullException(nameof(operation));

			try
			{
				// 显示开始提示
				progress?.Report(new AsyncProgress(0,$"开始{operationName}..."));

				// 执行异步操作
				await operation(progress,cancellationToken);

				// 完成提示
				progress?.Report(new AsyncProgress(100,$"{operationName}完成"));
				Toast.Show($"{operationName}完成",Toast.ToastType.Success);
			} catch(OperationCanceledException)
			{
				progress?.Report(new AsyncProgress(0,$"{operationName}已取消"));
				Toast.Show($"{operationName}已取消",Toast.ToastType.Info);
			} catch(Exception ex)
			{
				ExHandler.Run(() => throw ex,$"{operationName}执行失败");
				Toast.Show($"{operationName}失败: {ex.Message}",Toast.ToastType.Error);
			}
		}

		/// <summary>
		/// 执行异步操作（带返回值）
		/// </summary>
		/// <typeparam name="T"> 返回值类型 </typeparam>
		/// <param name="operation"> 要执行的操作 </param>
		/// <param name="progress"> 进度报告对象 </param>
		/// <param name="cancellationToken"> 取消令牌 </param>
		/// <param name="operationName"> 操作名称 </param>
		/// <returns> 操作结果 </returns>
		public static async Task<T> ExecuteAsync<T>(
			Func<IProgress<AsyncProgress>,CancellationToken,Task<T>> operation,
			IProgress<AsyncProgress> progress = null,
			CancellationToken cancellationToken = default,
			string operationName = "操作")
		{
			if(operation==null)
				throw new ArgumentNullException(nameof(operation));

			try
			{
				progress?.Report(new AsyncProgress(0,$"开始{operationName}..."));
				var result = await operation(progress,cancellationToken);
				progress?.Report(new AsyncProgress(100,$"{operationName}完成"));
				Toast.Show($"{operationName}完成",Toast.ToastType.Success);
				return result;
			} catch(OperationCanceledException)
			{
				progress?.Report(new AsyncProgress(0,$"{operationName}已取消"));
				Toast.Show($"{operationName}已取消",Toast.ToastType.Info);
				throw;
			} catch(Exception ex)
			{
				ExHandler.Run(() => throw ex,$"{operationName}执行失败");
				Toast.Show($"{operationName}失败: {ex.Message}",Toast.ToastType.Error);
				throw;
			}
		}

		/// <summary>
		/// 执行异步操作，提供统一的异常处理和进度报告
		/// </summary>
		/// <param name="operation"> 要执行的异步操作 </param>
		/// <param name="operationName"> 操作名称（用于日志和性能监控） </param>
		/// <remarks>
		/// 此方法提供：
		/// 1. 统一的异常处理（通过 ExHandler）
		/// 2. 性能监控（记录执行时间）
		/// 3. 进度报告支持
		/// 4. 取消支持
		///
		/// 注意：此方法使用 async void，适用于 fire-and-forget 场景（如 Ribbon 事件处理）。 所有异常已在内部处理，不会导致未处理的异常。
		/// </remarks>
		public static async void ExecuteAsyncOperation(
			Func<Task> operation,
			string operationName = "异步操作")
		{
			if(operation==null)
				throw new ArgumentNullException(nameof(operation));

			var sw = Stopwatch.StartNew();
			var opName = string.IsNullOrWhiteSpace(operationName) ? "异步操作" : operationName;

			try
			{
				await operation();
				sw.Stop();
				Profiler.LogMessage(
					message: $"执行耗时: {sw.Elapsed.TotalMilliseconds:F3} ms",
					logLevel: "性能监控",
					callerMethod: opName,
					filePath: string.Empty);
			} catch(OperationCanceledException)
			{
				sw.Stop();
				// 用户取消，静默处理
				Profiler.LogMessage($"{opName}已取消 ({sw.Elapsed.TotalMilliseconds:F0}ms)","INFO");
			} catch(Exception ex)
			{
				sw.Stop();
				// 异常时记录详细信息并交由 ExHandler 处理
				Profiler.LogMessage($"{opName}失败 ({sw.Elapsed.TotalMilliseconds:F0}ms): {ex.GetType().Name}","ERROR");
				ExHandler.Run(() => throw ex,$"{opName}执行失败");
			}
		}
	}

	/// <summary>
	/// 异步操作进度报告
	/// </summary>
	public class AsyncProgress(int percentage,string message,int currentItem = 0,int totalItems = 0)
	{
		/// <summary>
		/// 进度百分比 (0-100)
		/// </summary>
		public int Percentage { get; } = Math.Max(0,Math.Min(100,percentage));

		/// <summary>
		/// 进度消息
		/// </summary>
		public string Message { get; } = message??string.Empty;

		/// <summary>
		/// 当前处理项索引
		/// </summary>
		public int CurrentItem { get; } = currentItem;

		/// <summary>
		/// 总项数
		/// </summary>
		public int TotalItems { get; } = totalItems;

		public override string ToString()
		{
			if(TotalItems>0)
				return $"{Message} ({CurrentItem}/{TotalItems})";
			return $"{Message} ({Percentage}%)";
		}
	}

	/// <summary>
	/// 进度指示器 - 使用 Toast 显示进度
	/// </summary>
	public class ProgressIndicator(string operationName):IProgress<AsyncProgress>
	{
		private readonly string _operationName = operationName??"操作";
		private int _lastPercentage = -1;
		private readonly object _lockObject = new();

		public void Report(AsyncProgress value)
		{
			lock(_lockObject)
			{
				// 避免频繁更新（每 10% 或每次项变化时更新）
				bool shouldUpdate = false;

				if(value.TotalItems>0)
				{
					// 有项数变化时更新
					shouldUpdate=true;
				} else if(value.Percentage/10!=_lastPercentage/10)
				{
					// 每 10% 更新一次
					shouldUpdate=true;
				}

				if(shouldUpdate)
				{
					_lastPercentage=value.Percentage;

					// 使用 Toast 显示进度（仅在关键节点显示）
					if(value.Percentage==0||value.Percentage==100||value.TotalItems>0)
					{
						string message = value.TotalItems > 0
								? $"{_operationName}: {value.CurrentItem}/{value.TotalItems}"
								: $"{_operationName}: {value.Percentage}%";

						Toast.Show(message,Toast.ToastType.Info,duration: 1000);
					}
					// 记录详细进度
					Profiler.LogMessage($"[进度] {_operationName} - {value}");
				}
			}
		}
	}
}

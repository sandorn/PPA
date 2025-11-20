using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction.Infrastructure;
using System;

namespace PPA.Core.Logging
{
	/// <summary>
	/// 帮助在没有显式注入的地方解析 ILogger。
	/// </summary>
	internal static class LoggerProvider
	{
		private static readonly ILogger _fallbackLogger = new ProfilerLoggerAdapter();

		public static ILogger GetLogger()
		{
			try
			{
				var provider = ApplicationProvider.Current?.ServiceProvider as IServiceProvider;
				if(provider!=null)
				{
					var logger = provider.GetService<ILogger>();
					if(logger!=null)
					{
						return logger;
					}
				}
			} catch
			{
				// 忽略解析异常，使用后备 Logger
			}

			return _fallbackLogger;
		}
	}
}

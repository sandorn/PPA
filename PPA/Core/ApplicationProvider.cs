using PPA.Core.Abstraction.Infrastructure;
using System;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core
{
	/// <summary>
	/// 默认的应用程序提供者，实现了对当前 PowerPoint 上下文的封装。
	/// </summary>
	public sealed class ApplicationProvider:IApplicationProvider
	{
		private static ApplicationProvider _current;

		public ApplicationProvider()
		{
			_current=this;
		}

		public static IApplicationProvider Current => _current;

		public NETOP.Application NetApplication { get; private set; }

		public MSOP.Application NativeApplication { get; private set; }

		public IServiceProvider ServiceProvider { get; private set; }

		public bool IsInitialized => NetApplication!=null||NativeApplication!=null;

		/// <summary>
		/// 更新当前上下文信息，由 ThisAddIn 在生命周期内调用。
		/// </summary>
		public void SetContext(NETOP.Application netApp,MSOP.Application nativeApp,IServiceProvider serviceProvider)
		{
			NetApplication=netApp;
			NativeApplication=nativeApp;
			ServiceProvider=serviceProvider;
		}
	}
}

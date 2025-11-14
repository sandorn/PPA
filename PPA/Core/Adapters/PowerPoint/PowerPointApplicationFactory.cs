using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters.PowerPoint
{
	/// <summary>
	/// PowerPoint 平台应用工厂
	/// </summary>
	public sealed class PowerPointApplicationFactory : IApplicationFactory
	{
		public ApplicationType CurrentType => ApplicationType.PowerPoint;

		public IApplication GetCurrent()
		{
			// 优先使用已初始化的 NetOffice 实例
			var netApp = Globals.ThisAddIn?.NetApp;
			if(netApp!=null)
				return new PowerPointApplication(netApp);

			// 尝试从原生 Application 包装
			var native = Globals.ThisAddIn?.NativeApp;
			if(native!=null)
			{
				try
				{
					var wrapped = new NETOP.Application(null,native);
					return new PowerPointApplication(wrapped);
				} catch
				{
					return null;
				}
			}

			return null;
		}
	}
}



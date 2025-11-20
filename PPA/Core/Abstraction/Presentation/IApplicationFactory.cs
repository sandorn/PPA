namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 应用程序工厂接口 根据运行环境提供 IApplication 实例
	/// </summary>
	/// <remarks>
	/// 此接口定义了应用程序工厂的接口，用于根据当前运行环境创建相应的应用程序对象。 注意：当前版本仅支持 PowerPoint，WPS 支持已废弃。 实现类： <see cref="PPA.Core.Adapters.PowerPoint.PowerPointApplicationFactory" />。
	/// </remarks>
	public interface IApplicationFactory
	{
		/// <summary>
		/// 获取当前平台类型
		/// </summary>
		/// <value> 应用程序类型枚举值，参见 <see cref="ApplicationType" /> </value>
		/// <remarks> 此属性返回当前检测到的应用程序类型。 如果无法检测或未初始化，则返回 <see cref="ApplicationType.Unknown" />。 </remarks>
		ApplicationType CurrentType { get; }

		/// <summary>
		/// 获取当前平台的应用抽象对象
		/// </summary>
		/// <returns> 应用程序对象，如果无法探测或未初始化则返回 null </returns>
		/// <remarks> 此方法会根据当前运行环境创建相应的应用程序对象。 如果当前环境不是支持的应用程序（如 PowerPoint），则返回 null。 </remarks>
		IApplication GetCurrent();
	}
}

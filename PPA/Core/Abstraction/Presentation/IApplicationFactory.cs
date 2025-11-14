namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 跨平台应用工厂：根据运行环境提供 IApplication 实例
	/// </summary>
	public interface IApplicationFactory
	{
		/// <summary>
		/// 当前平台类型
		/// </summary>
		ApplicationType CurrentType { get; }

		/// <summary>
		/// 获取当前平台的应用抽象
		/// 返回 null 表示无法探测或未初始化
		/// </summary>
		IApplication GetCurrent();
	}
}



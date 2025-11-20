namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// Ribbon XML 提供者接口 负责加载和提供 Ribbon 功能区的 XML 配置
	/// </summary>
	/// <remarks> 此接口定义了 Ribbon XML 加载的接口，通过依赖注入使用，便于测试和扩展。 实现类可以从多种来源加载 XML，如嵌入式资源、文件系统等。 </remarks>
	public interface IRibbonXmlProvider
	{
		/// <summary>
		/// 获取 Ribbon XML 字符串
		/// </summary>
		/// <param name="ribbonID"> 功能区标识符，用于标识不同的功能区，通常为 "Microsoft.PowerPoint.Presentation" </param>
		/// <returns> Ribbon 的 XML 配置字符串，如果加载失败则返回 null 或空字符串 </returns>
		/// <remarks> 此方法会根据 ribbonID 加载对应的 Ribbon XML 配置。 XML 配置定义了功能区的布局、按钮、菜单等 UI 元素。 </remarks>
		string GetRibbonXml(string ribbonID);
	}
}

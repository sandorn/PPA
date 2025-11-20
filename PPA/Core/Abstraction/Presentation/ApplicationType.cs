namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 支持的演示文稿应用程序类型枚举
	/// </summary>
	/// <remarks> 此枚举定义了项目支持的演示文稿应用程序类型。 注意：当前版本仅支持 PowerPoint，WPS 支持已废弃。 </remarks>
	public enum ApplicationType
	{
		/// <summary>
		/// 未知类型或无法检测的应用程序
		/// </summary>
		Unknown = 0,

		/// <summary>
		/// Microsoft PowerPoint 应用程序
		/// </summary>
		PowerPoint = 1,

		/// <summary>
		/// WPS 演示应用程序（已废弃，当前版本不支持 WPS）
		/// </summary>
		[System.Obsolete("WPS 支持已废弃，当前版本仅支持 PowerPoint", false)]
		WpsPresentation = 2
	}
}

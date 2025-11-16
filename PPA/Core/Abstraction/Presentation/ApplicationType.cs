namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 支持的演示文稿应用程序类型
	/// </summary>
	public enum ApplicationType
	{
		Unknown = 0,
		PowerPoint = 1,
		/// <summary>
		/// WPS 演示（已废弃，当前版本不支持 WPS）
		/// </summary>
		[System.Obsolete("WPS 支持已废弃，当前版本仅支持 PowerPoint", false)]
		WpsPresentation = 2
	}
}



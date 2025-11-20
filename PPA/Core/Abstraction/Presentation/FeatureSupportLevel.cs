namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 功能支持级别枚举 用于描述不同平台对特性的支持情况
	/// </summary>
	/// <remarks>
	/// 此枚举用于 <see cref="IApplication.GetFeatureSupport(string)" /> 方法， 用于检查当前应用程序是否支持指定的功能。 可以用于实现功能降级或提示用户当前环境不支持某些功能。
	/// </remarks>
	public enum FeatureSupportLevel
	{
		/// <summary>
		/// 不支持该功能 当前平台完全不支持该功能，无法使用
		/// </summary>
		Unsupported = 0,

		/// <summary>
		/// 部分支持该功能 当前平台支持该功能的部分特性，但可能有一些限制或差异
		/// </summary>
		Partial = 1,

		/// <summary>
		/// 完全支持该功能 当前平台完全支持该功能，所有特性都可用
		/// </summary>
		Full = 2
	}
}

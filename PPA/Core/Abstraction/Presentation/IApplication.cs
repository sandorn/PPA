using System;
using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示演示文稿应用程序的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口主要用于：
	/// 1. 依赖注入（DI）集成：通过 IApplicationFactory 统一获取应用程序对象
	/// 2. 单元测试支持：可以创建 Mock 对象进行测试
	/// 3. 代码解耦：业务层不直接依赖具体的 NetOffice 类型
	/// 
	/// 注意：当前版本仅支持 PowerPoint，WPS 支持已废弃。
	/// 抽象接口的设计不是为了多平台支持，而是为了测试和 DI 的便利性。
	/// </remarks>
	public interface IApplication : IComWrapper
	{
		/// <summary>
		/// 获取应用程序类型（当前仅支持 PowerPoint）
		/// </summary>
		ApplicationType ApplicationType { get; }

		/// <summary>
		/// 获取当前激活的演示文稿，没有时返回 null
		/// </summary>
		IPresentation GetActivePresentation();

		/// <summary>
		/// 获取当前激活的幻灯片，没有时返回 null
		/// </summary>
		ISlide GetActiveSlide();

		/// <summary>
		/// 获取用户当前的选中对象集合（可能为空集合）
		/// </summary>
		IReadOnlyList<IShape> GetSelectedShapes();

		/// <summary>
		/// 判断指定功能的支持程度
		/// </summary>
		/// <param name="featureKey">功能标识，例如 "TableFormatting"</param>
		/// <returns>功能支持级别</returns>
		FeatureSupportLevel GetFeatureSupport(string featureKey);

		/// <summary>
		/// 以统一方式执行需要 UI 线程的操作
		/// </summary>
		/// <param name="action">需要执行的操作</param>
		void RunOnUiThread(Action action);
	}
}



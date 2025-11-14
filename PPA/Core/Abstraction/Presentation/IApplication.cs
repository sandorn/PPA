using System;
using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示演示文稿应用程序的抽象接口
	/// </summary>
	public interface IApplication : IComWrapper
	{
		/// <summary>
		/// 获取应用程序类型（PowerPoint、WPS 等）
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



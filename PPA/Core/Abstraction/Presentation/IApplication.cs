using System;
using System.Collections.Generic;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// 表示演示文稿应用程序的抽象接口
	/// </summary>
	/// <remarks>
	/// 此接口主要用于：
	/// <list type="number">
	/// <item>
	/// <description> 依赖注入（DI）集成：通过 <see cref="IApplicationFactory" /> 统一获取应用程序对象 </description>
	/// </item>
	/// <item>
	/// <description> 单元测试支持：可以创建 Mock 对象进行测试 </description>
	/// </item>
	/// <item>
	/// <description> 代码解耦：业务层不直接依赖具体的 NetOffice 类型 </description>
	/// </item>
	/// </list>
	/// 注意：当前版本仅支持 PowerPoint，WPS 支持已废弃。 抽象接口的设计不是为了多平台支持，而是为了测试和 DI 的便利性。
	/// </remarks>
	public interface IApplication:IComWrapper
	{
		/// <summary>
		/// 获取应用程序类型（当前仅支持 PowerPoint）
		/// </summary>
		/// <value> 应用程序类型枚举值，参见 <see cref="ApplicationType" /> </value>
		ApplicationType ApplicationType { get; }

		/// <summary>
		/// 获取当前激活的演示文稿
		/// </summary>
		/// <returns> 当前激活的演示文稿对象，如果没有打开的演示文稿则返回 null </returns>
		/// <remarks> 此方法返回当前 PowerPoint 窗口中激活的演示文稿。 如果用户没有打开任何演示文稿，或者当前没有激活的窗口，则返回 null。 </remarks>
		IPresentation GetActivePresentation();

		/// <summary>
		/// 获取当前激活的幻灯片
		/// </summary>
		/// <returns> 当前激活的幻灯片对象，如果没有激活的幻灯片则返回 null </returns>
		/// <remarks> 此方法返回当前视图中的活动幻灯片。 如果当前没有激活的演示文稿或幻灯片，则返回 null。 </remarks>
		ISlide GetActiveSlide();

		/// <summary>
		/// 获取用户当前选中的形状集合
		/// </summary>
		/// <returns> 选中的形状集合，如果没有选中任何形状则返回空集合（不会返回 null） </returns>
		/// <remarks> 此方法返回当前幻灯片中用户选中的所有形状。 如果用户没有选中任何形状，则返回空集合（Count = 0）。 </remarks>
		IReadOnlyList<IShape> GetSelectedShapes();

		/// <summary>
		/// 判断指定功能的支持程度
		/// </summary>
		/// <param name="featureKey"> 功能标识，例如 "TableFormatting"、"ChartFormatting" 等 </param>
		/// <returns> 功能支持级别，参见 <see cref="FeatureSupportLevel" /> 枚举 </returns>
		/// <remarks> 此方法用于检查当前应用程序是否支持指定的功能。 可以用于实现功能降级或提示用户当前环境不支持某些功能。 </remarks>
		FeatureSupportLevel GetFeatureSupport(string featureKey);

		/// <summary>
		/// 以统一方式执行需要 UI 线程的操作
		/// </summary>
		/// <param name="action"> 需要在 UI 线程上执行的操作，不能为 null </param>
		/// <remarks>
		/// 此方法确保 action 在 UI 线程上执行，这对于需要访问 COM 对象的操作非常重要。 如果当前已经在 UI 线程上，则直接执行；否则会切换到 UI 线程执行。
		/// </remarks>
		void RunOnUiThread(Action action);
	}
}

using System;
using MSOP = Microsoft.Office.Interop.PowerPoint;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Infrastructure
{
	/// <summary>
	/// 抽象的应用程序提供者接口 用于向非 VSTO 代码暴露当前的 PowerPoint 上下文，消除静态依赖
	/// </summary>
	/// <remarks>
	/// 此接口提供了统一的应用程序上下文访问方式，包括：
	/// <list type="bullet">
	/// <item>
	/// <description> NetOffice Application 对象 - <see cref="NetApplication" /> </description>
	/// </item>
	/// <item>
	/// <description> 原生 COM Application 对象 - <see cref="NativeApplication" /> </description>
	/// </item>
	/// <item>
	/// <description> 依赖注入服务提供者 - <see cref="ServiceProvider" /> </description>
	/// </item>
	/// </list>
	/// 通过此接口，业务代码可以获取应用程序实例和服务，而无需直接依赖 VSTO 的 <c> Globals.ThisAddIn </c>。
	/// </remarks>
	public interface IApplicationProvider
	{
		/// <summary>
		/// 获取 NetOffice Application 实例
		/// </summary>
		/// <value> NetOffice 包装的 PowerPoint Application 对象，如果未初始化则可能为 null </value>
		/// <remarks> NetOffice 提供了更友好的 API 和更好的异常处理，推荐优先使用此属性。 </remarks>
		NETOP.Application NetApplication { get; }

		/// <summary>
		/// 获取原生 COM Application 实例
		/// </summary>
		/// <value>
		/// 原生 COM PowerPoint Application
		/// 对象（Microsoft.Office.Interop.PowerPoint.Application），如果未初始化则可能为 null
		/// </value>
		/// <remarks> 仅在需要直接访问底层 COM 接口时使用，大多数场景应使用 <see cref="NetApplication" />。 </remarks>
		MSOP.Application NativeApplication { get; }

		/// <summary>
		/// 获取当前的 DI 服务提供者
		/// </summary>
		/// <value> 依赖注入服务提供者，用于解析注册的服务，如果未初始化则可能为 null </value>
		/// <remarks>
		/// 使用此属性可以获取通过依赖注入注册的服务，如 <see cref="ILogger" />、 <see cref="IFormattingConfig" /> 等。
		/// </remarks>
		IServiceProvider ServiceProvider { get; }

		/// <summary>
		/// 获取当前上下文是否已初始化
		/// </summary>
		/// <value> 如果应用程序上下文已初始化则为 true，否则为 false </value>
		/// <remarks>
		/// 在访问 <see cref="NetApplication" />、 <see cref="NativeApplication" /> 或
		/// <see cref="ServiceProvider" /> 之前， 应检查此属性以确保上下文已正确初始化。
		/// </remarks>
		bool IsInitialized { get; }
	}
}

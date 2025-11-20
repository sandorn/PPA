using System.Drawing;
using Office = Microsoft.Office.Core;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// Ribbon 图标提供者接口 负责管理和提供 Ribbon 功能区图标资源
	/// </summary>
	/// <remarks> 此接口定义了 Ribbon 图标管理的接口，通过依赖注入使用，便于测试和扩展。 实现类负责从资源文件加载图标，并提供图标缓存机制以提升性能。 </remarks>
	public interface IRibbonIconProvider
	{
		/// <summary>
		/// 获取指定控件的图标
		/// </summary>
		/// <param name="control"> 功能区控件对象，不能为 null </param>
		/// <param name="pressed"> 切换按钮的按下状态（仅对切换按钮有效）。true 表示按下状态图标，false 表示未按下状态图标，null 表示默认图标 </param>
		/// <returns> 图标位图，如果未找到则返回 null </returns>
		/// <remarks> 此方法会首先从缓存中查找图标，如果缓存中没有，则从资源文件加载并缓存。 对于切换按钮，可以根据 pressed 参数返回不同状态的图标。 </remarks>
		Bitmap GetIcon(Office.IRibbonControl control,bool? pressed = null);

		/// <summary>
		/// 预加载所有图标到缓存中
		/// </summary>
		/// <remarks> 此方法会在插件启动时调用，将所有图标预加载到内存中，以提升运行时性能。 建议在 Ribbon 加载前调用此方法。 </remarks>
		void PreloadIcons();

		/// <summary>
		/// 释放所有缓存的图标资源
		/// </summary>
		/// <remarks>
		/// 此方法会在插件关闭时调用，释放所有缓存的图标资源，避免内存泄漏。 调用此方法后，所有缓存的图标将被清空，后续调用 <see cref="GetIcon" /> 会重新加载。
		/// </remarks>
		void DisposeIcons();
	}
}

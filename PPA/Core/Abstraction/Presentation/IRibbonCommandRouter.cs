using Office = Microsoft.Office.Core;

namespace PPA.Core.Abstraction.Presentation
{
	/// <summary>
	/// Ribbon 命令路由接口 负责将按钮 ID 映射到相应的业务逻辑，实现 Ribbon UI 与业务逻辑的解耦
	/// </summary>
	/// <remarks>
	/// 此接口定义了 Ribbon 命令路由的接口，通过依赖注入使用，便于测试和扩展。 实现类负责处理所有 Ribbon 按钮点击事件，并根据按钮 ID 路由到相应的业务逻辑。
	/// </remarks>
	public interface IRibbonCommandRouter
	{
		/// <summary>
		/// 执行按钮命令
		/// </summary>
		/// <param name="buttonId"> 按钮标识符，例如 "Bt101", "Bt401" 等，不能为 null 或空字符串 </param>
		/// <returns> 如果命令成功执行则为 true，否则为 false </returns>
		/// <remarks>
		/// 此方法会根据 buttonId 路由到相应的业务逻辑，如格式化表格、文本、图表等操作。 如果 buttonId 未识别，会记录警告日志并返回 false。
		/// </remarks>
		bool ExecuteButtonCommand(string buttonId);

		/// <summary>
		/// 处理切换按钮的点击事件
		/// </summary>
		/// <param name="control"> 功能区控件对象，不能为 null </param>
		/// <param name="pressed"> 切换按钮的按下状态，true 表示按下，false 表示未按下 </param>
		/// <returns> 如果事件成功处理则为 true，否则为 false </returns>
		/// <remarks> 此方法处理切换按钮（ToggleButton）的点击事件，例如对齐基准切换按钮（Tb101）。 </remarks>
		bool HandleToggleButton(Office.IRibbonControl control,bool pressed);

		/// <summary>
		/// 处理菜单项的点击事件
		/// </summary>
		/// <param name="control"> 功能区控件对象，不能为 null </param>
		/// <returns> 如果事件成功处理则为 true，否则为 false </returns>
		/// <remarks> 此方法处理菜单项（MenuItem）的点击事件，例如设置菜单、关于菜单等。 </remarks>
		bool HandleMenuAction(Office.IRibbonControl control);
	}
}

using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 对齐类型枚举
	/// </summary>
	public enum AlignmentType
	{
		Left, Right, Top, Bottom, Centers, Middles, Horizontally, Vertically
	}

	/// <summary>
	/// 对齐工具辅助接口 提供形状对齐、分布、拉伸等相关操作
	/// </summary>
	/// <remarks>
	/// 此接口定义了形状对齐操作的接口，支持多种对齐类型和模式。 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口（
	/// <see cref="IApplication" />）。 实现类还提供其他方法（如 AttachLeft、SetEqualWidth 等），这些方法可以通过实例直接调用，但不强制在接口中定义。
	/// </remarks>
	public interface IAlignHelper
	{
		/// <summary>
		/// 执行对齐操作（NetOffice 版本）
		/// </summary>
		/// <param name="netApp"> NetOffice PowerPoint 应用程序实例，不能为 null </param>
		/// <param name="alignment"> 对齐类型，参见 <see cref="AlignmentType" /> 枚举 </param>
		/// <param name="alignToSlideMode"> 是否对齐到幻灯片。true 表示对齐到幻灯片边界，false 表示对齐到所选对象 </param>
		/// <remarks>
		/// 对齐行为：
		/// <list type="bullet">
		/// <item>
		/// <description> 单选形状：总是对齐到幻灯片 </description>
		/// </item>
		/// <item>
		/// <description> 多选形状：根据 alignToSlideMode 参数决定对齐基准 </description>
		/// </item>
		/// </list>
		/// </remarks>
		void ExecuteAlignment(NETOP.Application netApp,AlignmentType alignment,bool alignToSlideMode);

		/// <summary>
		/// 执行对齐操作（抽象接口版本）
		/// </summary>
		/// <param name="abstractApp"> 抽象应用程序实例，不能为 null </param>
		/// <param name="alignment"> 对齐类型，参见 <see cref="AlignmentType" /> 枚举 </param>
		/// <param name="alignToSlideMode"> 是否对齐到幻灯片。true 表示对齐到幻灯片边界，false 表示对齐到所选对象 </param>
		/// <remarks> 此方法内部会将抽象接口转换为 NetOffice 对象，然后调用 NetOffice 版本的方法。 </remarks>
		void ExecuteAlignment(IApplication abstractApp,AlignmentType alignment,bool alignToSlideMode);

		// 注意：其他方法（AttachLeft, SetEqualWidth 等）保持为公共方法，但不强制在接口中定义 这些方法可以通过实例直接调用
	}
}

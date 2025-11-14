using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;

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
	/// 对齐工具辅助接口
	/// 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface IAlignHelper
	{
		/// <summary>
		/// 执行对齐操作
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <param name="alignment">对齐类型</param>
		/// <param name="alignToSlideMode">是否对齐到幻灯片</param>
		void ExecuteAlignment(NETOP.Application app, AlignmentType alignment, bool alignToSlideMode);

		/// <summary>
		/// 执行对齐操作（抽象接口版本）
		/// </summary>
		/// <param name="app">抽象应用程序实例</param>
		/// <param name="alignment">对齐类型</param>
		/// <param name="alignToSlideMode">是否对齐到幻灯片</param>
		void ExecuteAlignment(IApplication app, AlignmentType alignment, bool alignToSlideMode);

		// 注意：其他方法（AttachLeft, SetEqualWidth 等）保持为公共方法，但不强制在接口中定义
		// 这些方法可以通过实例直接调用
	}
}


using PPA.Core.Abstraction.Presentation;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 形状工具辅助接口
	/// 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface IShapeHelper
	{
		/// <summary>
		/// 创建单个矩形
		/// </summary>
		/// <param name="slide">幻灯片对象</param>
		/// <param name="left">左边距</param>
		/// <param name="top">上边距</param>
		/// <param name="width">宽度</param>
		/// <param name="height">高度</param>
		/// <param name="rotation">旋转角度（可选）</param>
		/// <returns>创建的形状对象</returns>
		NETOP.Shape AddOneShape(NETOP.Slide slide, float left, float top, float width, float height, float rotation = 0);

		/// <summary>
		/// 获取形状的边框宽度
		/// </summary>
		/// <param name="shape">形状对象</param>
		/// <returns>边框宽度（上、左、右、下）</returns>
		(float top, float left, float right, float bottom) GetShapeBorderWeights(NETOP.Shape shape);

		/// <summary>
		/// 检查 COM 对象是否无效
		/// </summary>
		/// <param name="comObj">COM 对象</param>
		/// <returns>如果对象无效返回 true</returns>
		bool IsInvalidComObject(object comObj);

		/// <summary>
		/// 尝试获取当前幻灯片
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <returns>当前幻灯片对象，如果获取失败则返回 null</returns>
		NETOP.Slide TryGetCurrentSlide(NETOP.Application app);

		/// <summary>
		/// 验证并返回当前选择的对象
		/// </summary>
		/// <param name="app">PowerPoint 应用程序实例</param>
		/// <param name="requireMultipleShapes">是否要求必须选择多个形状</param>
		/// <returns>选择的对象（ShapeRange、Shape 或 null）</returns>
		dynamic ValidateSelection(NETOP.Application app, bool requireMultipleShapes = false);

		/// <summary>
		/// 尝试获取当前幻灯片（抽象接口版本）
		/// </summary>
		ISlide TryGetCurrentSlide(IApplication app);

		/// <summary>
		/// 验证并返回当前选择的对象（抽象接口版本）
		/// </summary>
		object ValidateSelection(IApplication app, bool requireMultipleShapes = false);
	}
}


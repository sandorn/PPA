using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 形状批量操作辅助接口 提供形状的批量操作功能，如创建外框、切换可见性等
	/// </summary>
	/// <remarks> 此接口定义了形状批量操作的常用功能，通过依赖注入使用，便于测试和扩展。 注意：当前方法签名使用 <see cref="NETOP.Application" />，后续可能改为使用平台抽象接口。 </remarks>
	public interface IShapeBatchHelper
	{
		/// <summary>
		/// 创建矩形外框（Bt601） 为选中的形状创建一个包围所有形状的矩形外框
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序实例，不能为 null </param>
		/// <remarks> 如果没有选中形状，则不会执行任何操作。 如果选中了多个形状，会创建一个包围所有形状的矩形。 </remarks>
		void CreateBoundingBox(NETOP.Application netApp);

		/// <summary>
		/// 切换形状可见性（Bt401） 如果选中了形状，则隐藏/显示选中的形状；如果没有选中，则显示所有隐藏的形状
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序实例，不能为 null </param>
		/// <remarks>
		/// 行为说明：
		/// <list type="bullet">
		/// <item>
		/// <description> 如果选中了形状，则切换选中形状的可见性 </description>
		/// </item>
		/// <item>
		/// <description> 如果没有选中形状，则显示当前幻灯片上所有隐藏的形状 </description>
		/// </item>
		/// </list>
		/// </remarks>
		void ToggleShapeVisibility(NETOP.Application netApp);
	}
}

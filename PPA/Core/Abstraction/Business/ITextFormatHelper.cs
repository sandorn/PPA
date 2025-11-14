using NETOP = NetOffice.PowerPointApi;
using PPA.Core.Abstraction.Presentation;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// 文本格式化辅助接口
	/// 注意：当前使用 NetOffice 类型，后续阶段将改为使用平台抽象接口
	/// </summary>
	public interface ITextFormatHelper
	{
		/// <summary>
		/// 应用文本格式化到指定形状
		/// </summary>
		/// <param name="shp">要格式化的形状对象</param>
		void ApplyTextFormatting(NETOP.Shape shp);

		/// <summary>
		/// 应用文本格式化到指定形状（抽象接口版本）
		/// </summary>
		/// <param name="shape">要格式化的抽象形状对象</param>
		void ApplyTextFormatting(IShape shape);
	}
}


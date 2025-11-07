using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Shape;
using PPA.Utilities;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 文本批量操作辅助类
	/// </summary>
	public static class TextBatchHelper
	{
		/// <summary>
		/// 批量格式化文本
		/// </summary>
		/// <param name="app"> PowerPoint 应用程序实例 </param>
		public static void Bt502_Click(NETOP.Application app)
		{
			UndoHelper.BeginUndoEntry(app,UndoHelper.UndoNames.FormatText);

			ExHandler.Run(() =>
			{
				// 获取选中的形状
				var selection = ShapeUtils.ValidateSelection(app);

				// 如果没有选中对象，显示提示并返回
				if(selection==null)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				bool hasFormatted = false;

				// 处理单个形状的情况
				if(selection is NETOP.Shape shape)
				{
					if(shape.TextFrame?.HasText==MsoTriState.msoTrue)
					{
						TextFormatHelper.ApplyTextFormatting(shape);
						hasFormatted=true;
					}
				}
				// 处理多个形状的情况
				else if(selection is NETOP.ShapeRange shapeRange)
				{
					foreach(NETOP.Shape s in shapeRange)
					{
						if(s.TextFrame?.HasText==MsoTriState.msoTrue)
						{
							TextFormatHelper.ApplyTextFormatting(s);
							hasFormatted=true;
						}
					}
				}

				// 如果成功美化了文本，显示成功提示
				if(hasFormatted)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_Success"),Toast.ToastType.Success);
				} else
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoText"),Toast.ToastType.Warning);
				}
			});
		}
	}
}

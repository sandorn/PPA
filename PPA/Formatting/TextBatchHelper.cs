using NetOffice.OfficeApi.Enums;
using PPA.Core;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Adapters;
using PPA.Core.Logging;
using PPA.Formatting.Selection;
using PPA.Utilities;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Formatting
{
	/// <summary>
	/// 文本批量操作辅助类 提供文本批量格式化功能
	/// </summary>
	internal class TextBatchHelper(ITextFormatHelper textFormatHelper,IShapeHelper shapeHelper,ILogger logger = null):ITextBatchHelper
	{
		private readonly ITextFormatHelper _textFormatHelper = textFormatHelper??throw new ArgumentNullException(nameof(textFormatHelper));
		private readonly IShapeHelper _shapeHelper = shapeHelper??throw new ArgumentNullException(nameof(shapeHelper));
		private readonly ILogger _logger = logger??LoggerProvider.GetLogger();

		#region ITextBatchHelper 实现

		/// <summary>
		/// 格式化文本（同步方法）
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		public void FormatText(NETOP.Application netApp)
		{
			if(netApp==null) throw new ArgumentNullException(nameof(netApp));
			FormatTextInternal(netApp,_textFormatHelper);
		}

		#endregion ITextBatchHelper 实现

		#region 内部实现

		/// <summary>
		/// 格式化文本的内部实现（同步）
		/// </summary>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		/// <param name="textFormatHelper"> 文本格式化辅助类 </param>
		private void FormatTextInternal(NETOP.Application netApp,ITextFormatHelper textFormatHelper)
		{
			_logger.LogInformation($"启动，netApp类型={netApp?.GetType().Name??"null"}");
			if(textFormatHelper==null)
				throw new InvalidOperationException("无法获取 ITextFormatHelper 服务");

			UndoHelper.BeginUndoEntry(netApp,UndoHelper.UndoNames.FormatText);

			ExHandler.Run(() =>
			{
				var abstractApp = ApplicationHelper.GetAbstractApplication(netApp);
				var selection = _shapeHelper.ValidateSelection(abstractApp) as dynamic;

				// 调试：记录选中对象信息
				if(selection==null)
				{
					_logger.LogWarning("ValidateSelection 返回 null，没有选中对象");
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoSelection"),Toast.ToastType.Warning);
					return;
				}

				// 调试：记录选中对象的数量和类型
				try
				{
					if(selection is NETOP.ShapeRange shapeRange)
					{
						int count = ExHandler.SafeGet(() => shapeRange.Count, defaultValue: 0);
						_logger.LogInformation($"选中对象类型=ShapeRange, 数量={count}");
					} else if(selection is NETOP.Shape shape)
					{
						_logger.LogInformation("选中对象类型=Shape, 数量=1");
					} else
					{
						_logger.LogInformation($"选中对象类型={selection?.GetType().Name??"null"}");
					}
				} catch(Exception ex)
				{
					_logger.LogWarning($"获取选中对象信息失败: {ex.Message}");
				}

				// 处理选中的形状
				bool hasFormatted = ProcessTextShapesFromSelection(selection, netApp, textFormatHelper);

				// 显示结果
				if(hasFormatted)
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_Success"),Toast.ToastType.Success);
				} else
				{
					Toast.Show(ResourceManager.GetString("Toast_FormatText_NoText"),Toast.ToastType.Warning);
				}
			},enableTiming: true);
		}

		/// <summary>
		/// 检查形状是否包含表格
		/// </summary>
		/// <param name="shape"> 形状对象 </param>
		/// <returns> 如果包含表格返回 true，否则返回 false </returns>
		private bool HasTable(NETOP.Shape shape)
		{
			if(shape==null) return false;

			bool hasTable = ExHandler.SafeGet(() => shape.HasTable == MsoTriState.msoTrue, defaultValue: false);
			if(hasTable) return true;

			var table = ExHandler.SafeGet(() => shape.Table, defaultValue: (NETOP.Table)null);
			return table!=null;
		}

		/// <summary>
		/// 检查形状是否包含文本
		/// </summary>
		/// <param name="shape"> 形状对象 </param>
		/// <returns> 如果包含文本返回 true，否则返回 false </returns>
		private bool HasText(NETOP.Shape shape)
		{
			if(shape==null) return false;

			// 方式1：尝试通过 TextFrame.HasText 属性
			bool hasText = ExHandler.SafeGet(() => shape.TextFrame?.HasText == MsoTriState.msoTrue, defaultValue: false);
			if(hasText) return true;

			// 方式2：检查 TextFrame 是否存在且有文本内容
			var textFrame = ExHandler.SafeGet(() => shape.TextFrame, defaultValue: (NETOP.TextFrame)null);
			if(textFrame!=null)
			{
				var textRange = ExHandler.SafeGet(() => textFrame.TextRange, defaultValue: (NETOP.TextRange)null);
				if(textRange!=null)
				{
					string text = ExHandler.SafeGet(() => textRange.Text, defaultValue: null);
					return !string.IsNullOrWhiteSpace(text);
				}
			}

			return false;
		}

		/// <summary>
		/// 处理单个文本形状
		/// </summary>
		/// <param name="shape"> 形状对象 </param>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		/// <param name="textFormatHelper"> 文本格式化辅助类 </param>
		/// <returns> 如果成功格式化返回 true，否则返回 false </returns>
		private bool ProcessTextShape(NETOP.Shape shape,NETOP.Application netApp,ITextFormatHelper textFormatHelper)
		{
			if(shape==null) return false;

			// 跳过表格内的文本（表格有自己的格式化逻辑）
			if(HasTable(shape))
			{
				return false;
			}

			// 检查是否有文本
			if(!HasText(shape))
			{
				return false;
			}

			// 包装并格式化
			var iShape = AdapterUtils.WrapShape(netApp, shape);
			if(iShape!=null)
			{
				textFormatHelper.ApplyTextFormatting(iShape);
				return true;
			}

			return false;
		}

		/// <summary>
		/// 从选区处理文本形状
		/// </summary>
		/// <param name="selection"> 选区对象 </param>
		/// <param name="netApp"> PowerPoint 应用程序对象 </param>
		/// <param name="textFormatHelper"> 文本格式化辅助类 </param>
		/// <returns> 如果至少格式化了一个形状返回 true，否则返回 false </returns>
		private bool ProcessTextShapesFromSelection(object selection,NETOP.Application netApp,ITextFormatHelper textFormatHelper)
		{
			var shapeSelection = ShapeSelectionFactory.Create(selection);
			if(shapeSelection==null)
			{
				_logger.LogWarning("无法识别的选区类型，跳过处理");
				return false;
			}

			bool hasFormatted = false;
			foreach(var shape in shapeSelection)
			{
				if(ProcessTextShape(shape,netApp,textFormatHelper))
				{
					hasFormatted=true;
				}
			}

			return hasFormatted;
		}

		#endregion 内部实现
	}
}

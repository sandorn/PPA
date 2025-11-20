using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters.PowerPoint;
using PPA.Core.Logging;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Adapters
{
	/// <summary>
	/// 统一的抽象适配器工具类 提供 NetOffice 对象与抽象接口之间的双向转换功能，减少各处重复的包装代码
	/// </summary>
	/// <remarks>
	/// 此类实现了适配器模式，用于在 NetOffice 包装的 COM 对象和项目定义的抽象接口之间进行转换。 主要功能包括：
	/// - 将 NetOffice 对象包装为抽象接口（Wrap 方法）
	/// - 从抽象接口提取 NetOffice 对象（Unwrap 方法）
	/// - 提供统一的错误处理和日志记录
	/// </remarks>
	public static class AdapterUtils
	{
		private static readonly ILogger Logger = LoggerProvider.GetLogger();

		/// <summary>
		/// 将 NetOffice Application 对象包装为 IApplication 接口
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <returns> IApplication 接口实例，如果 app 为 null 则返回 null </returns>
		public static IApplication WrapApplication(NETOP.Application app)
		{
			if(app==null) return null;
			return new PowerPointApplication(app);
		}

		/// <summary>
		/// 从 Shape 对象包装为 ISlide 接口
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="shape"> NetOffice Shape 对象，用于获取其所在的 Slide </param>
		/// <returns> ISlide 接口实例，如果无法获取则返回 null </returns>
		public static ISlide WrapSlide(NETOP.Application app,NETOP.Shape shape)
		{
			if(app==null||shape==null) return null;
			var iApp = WrapApplication(app);
			NETOP.Slide nativeSlide = null;
			try { nativeSlide=shape.Parent as NETOP.Slide; } catch { /* ignore */ }
			if(nativeSlide==null) return null;
			var iPres = WrapPresentation(iApp,nativeSlide);
			return new PowerPointSlide(iApp,iPres,nativeSlide);
		}

		/// <summary>
		/// 将 NetOffice Slide 对象包装为 ISlide 接口
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="slide"> NetOffice Slide 对象 </param>
		/// <returns> ISlide 接口实例，如果 app 或 slide 为 null 则返回 null </returns>
		public static ISlide WrapSlide(NETOP.Application app,NETOP.Slide slide)
		{
			if(app==null||slide==null) return null;
			var iApp = WrapApplication(app);
			var iPres = WrapPresentation(iApp,slide);
			return new PowerPointSlide(iApp,iPres,slide);
		}

		/// <summary>
		/// 从 Slide 对象包装为 IPresentation 接口
		/// </summary>
		/// <param name="iApp"> 抽象 Application 接口 </param>
		/// <param name="nativeSlide"> NetOffice Slide 对象，用于获取其所在的 Presentation </param>
		/// <returns> IPresentation 接口实例，如果无法获取则返回 null </returns>
		/// <remarks>
		/// 此方法会尝试多种方式获取 Presentation：
		/// 1. 从 slide.Parent 获取（最直接的方式）
		/// 2. 如果失败，从 Application.ActivePresentation 获取（备用方式）
		/// </remarks>
		public static IPresentation WrapPresentation(IApplication iApp,NETOP.Slide nativeSlide)
		{
			if(iApp==null||nativeSlide==null)
			{
				Logger.LogWarning($"iApp 或 nativeSlide 为 null (iApp={iApp?.GetType().Name??"null"}, nativeSlide={nativeSlide?.GetType().Name??"null"})");
				return null;
			}

			// 先验证 nativeSlide 是否有效
			bool slideValid = ExHandler.SafeGet(() =>
			{
				var _ = nativeSlide.SlideIndex;
				return true;
			}, defaultValue: false);

			if(!slideValid)
			{
				Logger.LogWarning("nativeSlide 无效（无法访问 SlideIndex），尝试从 Application 获取");
			}

			NETOP.Presentation nativePres = null;

			// 如果 slide 有效，尝试从 slide.Parent 获取
			if(slideValid)
			{
				try
				{
					nativePres=ExHandler.SafeGet(() => nativeSlide.Parent as NETOP.Presentation,defaultValue: (NETOP.Presentation) null);
					if(nativePres!=null)
					{
						Logger.LogDebug("从 slide.Parent 获取到 Presentation");
					}
				} catch(Exception ex)
				{
					Logger.LogWarning($"从 slide.Parent 获取 Presentation 失败: {ex.Message}");
				}
			}

			// 如果无法从 slide.Parent 获取，尝试从 Application 获取
			if(nativePres==null&&iApp is PowerPointApplication pptApp)
			{
				try
				{
					var netApp = pptApp.NativeObject;
					if(netApp!=null)
					{
						nativePres=ExHandler.SafeGet(() => netApp.ActivePresentation,defaultValue: (NETOP.Presentation) null);
						if(nativePres!=null)
						{
							Logger.LogDebug("从 app.ActivePresentation 获取到 Presentation");
						} else
						{
							Logger.LogWarning("app.ActivePresentation 返回 null");
						}
					} else
					{
						Logger.LogWarning("pptApp.NativeObject 为 null");
					}
				} catch(Exception ex)
				{
					Logger.LogWarning($"从 app.ActivePresentation 获取失败: {ex.Message}");
				}
			} else if(nativePres==null&&!(iApp is PowerPointApplication))
			{
				Logger.LogWarning($"iApp 不是 PowerPointApplication 类型 ({iApp?.GetType().Name??"null"})，无法从 Application 获取");
			}

			if(nativePres==null)
			{
				Logger.LogError("所有方法都失败，返回 null");
			}

			return nativePres!=null ? new PowerPointPresentation(iApp,nativePres) : null;
		}

		/// <summary>
		/// 将 NetOffice Shape 对象包装为 IShape 接口（自动获取 Slide）
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="shape"> NetOffice Shape 对象 </param>
		/// <returns> IShape 接口实例，如果无法获取则返回 null </returns>
		public static IShape WrapShape(NETOP.Application app,NETOP.Shape shape)
		{
			return WrapShape(app,shape,null);
		}

		/// <summary>
		/// 将 NetOffice Shape 对象包装为 IShape 接口
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="shape"> NetOffice Shape 对象 </param>
		/// <param name="slide"> 可选的 NetOffice Slide 对象，如果为 null 则尝试自动获取 </param>
		/// <returns> IShape 接口实例，如果无法获取则返回 null </returns>
		/// <remarks>
		/// 如果未提供 slide 参数，此方法会尝试以下方式获取：
		/// 1. 从 shape.Parent 获取
		/// 2. 从 app.ActiveWindow.View.Slide 获取（备用方式）
		/// </remarks>
		public static IShape WrapShape(NETOP.Application app,NETOP.Shape shape,NETOP.Slide slide)
		{
			Logger.LogInformation($"启动，app类型={app?.GetType().Name??"null"}, shape类型={shape?.GetType().Name??"null"}, slide类型={slide?.GetType().Name??"null"}");
			if(app==null||shape==null)
			{
				Logger.LogWarning("app 或 shape 为 null，返回 null");
				return null;
			}

			var iApp = WrapApplication(app);
			NETOP.Slide nativeSlide = slide;

			// 如果未提供 slide，尝试获取
			if(nativeSlide==null)
			{
				try { nativeSlide=shape.Parent as NETOP.Slide; } catch { /* ignore */ }
			}

			// 如果仍然无法获取 Slide，尝试从当前活动窗口获取
			if(nativeSlide==null)
			{
				try
				{
					var activeWindow = app.ActiveWindow;
					if(activeWindow!=null&&activeWindow.View!=null)
					{
						nativeSlide=activeWindow.View.Slide as NETOP.Slide;
					}
				} catch { /* ignore */ }
			}

			// 如果仍然无法获取 Slide，返回 null，避免创建无效的对象
			if(nativeSlide==null)
			{
				Logger.LogWarning("无法获取 shape 所在的 Slide，返回 null");
				return null;
			}

			var iPres = WrapPresentation(iApp, nativeSlide);
			if(iPres==null)
			{
				Logger.LogWarning("无法获取 Presentation，返回 null");
				return null;
			}

			var iSlide = new PowerPointSlide(iApp, iPres, nativeSlide);
			return new PowerPointShape(iApp,iPres,iSlide,shape);
		}

		/// <summary>
		/// 将 NetOffice Table 对象包装为 ITable 接口（自动获取 Slide）
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="shape"> 包含表格的 NetOffice Shape 对象 </param>
		/// <param name="table"> NetOffice Table 对象 </param>
		/// <returns> ITable 接口实例，如果无法获取则返回 null </returns>
		public static ITable WrapTable(NETOP.Application app,NETOP.Shape shape,NETOP.Table table)
		{
			return WrapTable(app,shape,table,null);
		}

		/// <summary>
		/// 将 NetOffice Table 对象包装为 ITable 接口
		/// </summary>
		/// <param name="app"> NetOffice PowerPoint Application 对象 </param>
		/// <param name="shape"> 包含表格的 NetOffice Shape 对象 </param>
		/// <param name="table"> NetOffice Table 对象 </param>
		/// <param name="slide"> 可选的 NetOffice Slide 对象，如果为 null 则尝试自动获取 </param>
		/// <returns> ITable 接口实例，如果无法获取则返回 null </returns>
		public static ITable WrapTable(NETOP.Application app,NETOP.Shape shape,NETOP.Table table,NETOP.Slide slide)
		{
			Logger.LogInformation($"WrapTable 被调用，app类型={app?.GetType().Name??"null"}, shape类型={shape?.GetType().Name??"null"}, table类型={table?.GetType().Name??"null"}, slide类型={slide?.GetType().Name??"null"}");
			if(app==null||shape==null||table==null)
			{
				Logger.LogWarning("app、shape 或 table 为 null，返回 null");
				return null;
			}

			var iShape = WrapShape(app, shape, slide);
			if(iShape==null)
			{
				Logger.LogError("无法创建表格");
				return null;
			}
			return new PowerPointTable(iShape,table);
		}

		#region 反向转换方法（从抽象接口到 NetOffice）

		/// <summary>
		/// 从抽象接口获取 NetOffice Application 对象
		/// </summary>
		/// <param name="app"> 抽象 Application 接口 </param>
		/// <returns> NetOffice Application 对象，如果无法获取则返回 null </returns>
		public static NETOP.Application UnwrapApplication(IApplication app)
		{
			if(app==null) return null;

			if(app is IComWrapper<NETOP.Application> typed)
			{
				return typed.NativeObject;
			}

			if(app is IComWrapper wrapper)
			{
				return wrapper.NativeObject as NETOP.Application;
			}

			return null;
		}

		/// <summary>
		/// 从抽象接口获取 NetOffice Slide 对象
		/// </summary>
		/// <param name="slide"> 抽象 Slide 接口 </param>
		/// <returns> NetOffice Slide 对象，如果无法获取则返回 null </returns>
		public static NETOP.Slide UnwrapSlide(ISlide slide)
		{
			if(slide==null) return null;

			if(slide is IComWrapper<NETOP.Slide> typed)
			{
				return typed.NativeObject;
			}

			if(slide is IComWrapper wrapper)
			{
				return wrapper.NativeObject as NETOP.Slide;
			}

			return null;
		}

		/// <summary>
		/// 从抽象接口获取 NetOffice Shape 对象
		/// </summary>
		/// <param name="shape"> 抽象 Shape 接口 </param>
		/// <returns> NetOffice Shape 对象，如果无法获取则返回 null </returns>
		public static NETOP.Shape UnwrapShape(IShape shape)
		{
			if(shape==null) return null;

			if(shape is IComWrapper<NETOP.Shape> typed)
			{
				return typed.NativeObject;
			}

			if(shape is IComWrapper wrapper)
			{
				return wrapper.NativeObject as NETOP.Shape;
			}

			return null;
		}

		/// <summary>
		/// 从抽象接口获取 NetOffice Application 对象（如果 slide 不为 null，则从 slide.Application 获取）
		/// </summary>
		/// <param name="slide"> 抽象 Slide 接口 </param>
		/// <returns> NetOffice Application 对象，如果无法获取则返回 null </returns>
		public static NETOP.Application UnwrapApplicationFromSlide(ISlide slide)
		{
			if(slide==null) return null;

			var app = slide.Application;
			if(app==null) return null;

			return UnwrapApplication(app);
		}

		/// <summary>
		/// 从抽象接口获取 NetOffice Table 对象
		/// </summary>
		/// <param name="table"> 抽象 Table 接口 </param>
		/// <returns> NetOffice Table 对象，如果无法获取则返回 null </returns>
		public static NETOP.Table UnwrapTable(ITable table)
		{
			if(table==null) return null;

			if(table is IComWrapper<NETOP.Table> typed)
			{
				return typed.NativeObject;
			}

			if(table is IComWrapper wrapper)
			{
				return wrapper.NativeObject as NETOP.Table;
			}

			return null;
		}

		#endregion 反向转换方法（从抽象接口到 NetOffice）
	}
}

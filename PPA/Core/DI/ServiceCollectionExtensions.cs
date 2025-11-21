using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Logging;
using PPA.Manipulation;
using PPA.UI.Providers;
using PPA.Utilities;

namespace PPA.Core.DI
{
	/// <summary>
	/// DI 容器扩展方法 用于注册 PPA 项目的所有服务
	/// </summary>
	public static class ServiceCollectionExtensions
	{
		/// <summary>
		/// 添加 PPA 项目的所有服务到 DI 容器
		/// </summary>
		/// <param name="services"> 服务集合 </param>
		/// <returns> 服务集合，支持链式调用 </returns>
		public static IServiceCollection AddPPAServices(this IServiceCollection services)
		{
			// 注册配置服务（单例）
			services.AddSingleton<IFormattingConfig>(sp => FormattingConfig.Instance);

			// 注册应用程序提供者，统一暴露当前 Application 上下文
			services.AddSingleton<ApplicationProvider>();
			services.AddSingleton<IApplicationProvider>(sp => sp.GetRequiredService<ApplicationProvider>());

			// 注册默认日志实现
			services.AddSingleton<ILogger,ProfilerLoggerAdapter>();

			// 注册 Ribbon 相关服务
			services.AddSingleton<IRibbonXmlProvider,EmbeddedRibbonXmlProvider>();
			services.AddSingleton<IRibbonIconProvider,RibbonIconProvider>();
			// 注意：IRibbonCommandRouter 需要从 CustomRibbon 创建，因为它需要回调函数

			// 注册格式化辅助服务（瞬态，每次请求创建新实例）
			services.AddTransient<ISelectionService, SelectionService>();
			services.AddTransient<ITableFormatHelper,TableFormatHelper>();
			services.AddTransient<ITextFormatHelper,TextFormatHelper>();
			services.AddTransient<IChartFormatHelper,ChartFormatHelper>();
			services.AddTransient<IAlignHelper,AlignHelper>();
			services.AddTransient<ITableBatchHelper,TableBatchHelper>();
			services.AddTransient<ITextBatchHelper,TextBatchHelper>();
			services.AddTransient<IChartBatchHelper,ChartBatchHelper>();
			services.AddTransient<IShapeBatchHelper,ShapeBatchHelper>();

			// 注册工具服务（单例，因为是无状态的工具类）
			services.AddSingleton<IShapeHelper,PPA.Shape.ShapeUtils>();

			// 注册命令执行器（瞬态，需要应用程序实例）
			services.AddTransient<ICommandExecutor,CommandExecutor>();

			// 平台抽象与适配器（仅 PowerPoint）：注册工厂 + IApplication 解析 注意：当前版本仅支持 PowerPoint，WPS 支持已废弃
			// 抽象接口主要用于：1) 依赖注入集成 2) 单元测试支持 3) 代码解耦
			// services.AddSingleton<PowerPointApplicationFactory>();
			// services.AddSingleton<IApplicationFactory>(sp =>
			// {
			// 	var factories = new IApplicationFactory[]
			// 	{
			// 		sp.GetRequiredService<PowerPointApplicationFactory>()
			// 	};
			// 	return new CompositeApplicationFactory(factories);
			// });
			// services.AddTransient<IApplication>(sp => sp.GetRequiredService<IApplicationFactory>()?.GetCurrent());

			// 注意：其他业务服务（IBatchHelper 等）将在后续步骤注册

			return services;
		}
	}
}

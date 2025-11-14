using Microsoft.Extensions.DependencyInjection;
using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Presentation;
using PPA.Core.Adapters;
using PPA.Core.Adapters.PowerPoint;
// using PPA.Core.Adapters.Wps; // WPS 支持已禁用
using PPA.Formatting;

namespace PPA.Core.DI
{
	/// <summary>
	/// DI 容器扩展方法
	/// 用于注册 PPA 项目的所有服务
	/// </summary>
	public static class ServiceCollectionExtensions
	{
		/// <summary>
		/// 添加 PPA 项目的所有服务到 DI 容器
		/// </summary>
		/// <param name="services">服务集合</param>
		/// <returns>服务集合，支持链式调用</returns>
		public static IServiceCollection AddPPAServices(this IServiceCollection services)
		{
			// 注册配置服务（单例）
			services.AddSingleton<IFormattingConfig>(sp => FormattingConfig.Instance);

			// 注册格式化辅助服务（瞬态，每次请求创建新实例）
			services.AddTransient<ITableFormatHelper, TableFormatHelper>();
			services.AddTransient<ITextFormatHelper, TextFormatHelper>();
			services.AddTransient<IChartFormatHelper, ChartFormatHelper>();
			services.AddTransient<IAlignHelper, AlignHelper>();
			services.AddTransient<ITableBatchHelper, TableBatchHelper>();
			services.AddTransient<ITextBatchHelper, TextBatchHelper>();
			services.AddTransient<IChartBatchHelper, ChartBatchHelper>();

			// 注册工具服务（单例，因为是无状态的工具类）
			services.AddSingleton<IShapeHelper, PPA.Shape.ShapeUtils>();

			// 平台抽象与适配器（仅 PowerPoint）：注册工厂 + IApplication 解析
			// WPS 支持已禁用，如需启用请取消注释下面的 WpsApplicationFactory 注册
			// services.AddSingleton<WpsApplicationFactory>();
			services.AddSingleton<PowerPointApplicationFactory>();
			services.AddSingleton<IApplicationFactory>(sp =>
			{
				var factories = new IApplicationFactory[]
				{
					// sp.GetRequiredService<WpsApplicationFactory>(), // WPS 支持已禁用
					sp.GetRequiredService<PowerPointApplicationFactory>()
				};
				return new CompositeApplicationFactory(factories);
			});
			services.AddTransient<IApplication>(sp => sp.GetRequiredService<IApplicationFactory>()?.GetCurrent());

			// 注意：其他业务服务（IBatchHelper 等）将在后续步骤注册

			return services;
		}
	}
}


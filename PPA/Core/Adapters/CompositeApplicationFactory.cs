using PPA.Core.Abstraction.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PPA.Core.Adapters
{
	/// <summary>
	/// 组合应用工厂：按顺序尝试多个工厂，返回首个可用平台
	/// </summary>
	/// <remarks> 当前版本仅支持 PowerPoint，但保留此设计以便未来扩展支持其他平台（如 WPS）。 工厂会按顺序尝试所有注册的工厂，返回第一个成功获取的应用程序实例。 </remarks>
	public sealed class CompositeApplicationFactory:IApplicationFactory
	{
		private readonly IReadOnlyList<IApplicationFactory> _factories;
		private readonly object _syncRoot = new();
		private IApplication _cachedInstance;

		public CompositeApplicationFactory(IEnumerable<IApplicationFactory> factories)
		{
			_factories=(factories??Enumerable.Empty<IApplicationFactory>())
				.Where(f => f!=null)
				.ToList()
				.AsReadOnly();
		}

		public ApplicationType CurrentType => GetCurrent()?.ApplicationType??ApplicationType.Unknown;

		public IApplication GetCurrent()
		{
			if(_cachedInstance!=null) return _cachedInstance;

			lock(_syncRoot)
			{
				if(_cachedInstance!=null) return _cachedInstance;

				foreach(var factory in _factories)
				{
					try
					{
						var instance = factory.GetCurrent();
						if(instance!=null)
						{
							_cachedInstance=instance;
							break;
						}
					} catch(Exception ex)
					{
						Profiler.LogMessage($"组合工厂调用 {factory?.GetType().Name} 失败: {ex.Message}");
					}
				}
			}

			return _cachedInstance;
		}
	}
}

using System;
using System.Collections.Generic;
using System.Linq;
using PPA.Core;
using PPA.Core.Abstraction.Presentation;

namespace PPA.Core.Adapters
{
	/// <summary>
	/// 组合应用工厂：按顺序尝试多个工厂，返回首个可用平台
	/// </summary>
	public sealed class CompositeApplicationFactory : IApplicationFactory
	{
		private readonly IReadOnlyList<IApplicationFactory> _factories;
		private readonly object _syncRoot = new();
		private IApplication _cachedInstance;

		public CompositeApplicationFactory(IEnumerable<IApplicationFactory> factories)
		{
			_factories = (factories ?? Enumerable.Empty<IApplicationFactory>())
				.Where(f => f != null)
				.ToList()
				.AsReadOnly();
		}

		public ApplicationType CurrentType => GetCurrent()?.ApplicationType ?? ApplicationType.Unknown;

		public IApplication GetCurrent()
		{
			if(_cachedInstance != null) return _cachedInstance;

			lock(_syncRoot)
			{
				if(_cachedInstance != null) return _cachedInstance;

				foreach(var factory in _factories)
				{
					try
					{
						var instance = factory.GetCurrent();
						if(instance != null)
						{
							_cachedInstance = instance;
							break;
						}
					}
					catch(Exception ex)
					{
						Profiler.LogMessage($"组合工厂调用 {factory?.GetType().Name} 失败: {ex.Message}");
					}
				}
			}

			return _cachedInstance;
		}
	}
}



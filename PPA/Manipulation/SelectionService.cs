using PPA.Core.Abstraction.Business;
using PPA.Core.Abstraction.Infrastructure;
using PPA.Core.Logging;
using System;
using NETOP = NetOffice.PowerPointApi;

namespace PPA.Manipulation
{
    public class SelectionService : ISelectionService
    {
        private readonly IApplicationProvider _applicationProvider;
        private readonly ILogger _logger;

        public SelectionService(IApplicationProvider applicationProvider, ILogger logger)
        {
            _applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            _logger = logger ?? LoggerProvider.GetLogger();
        }

        private NETOP.Application GetNetApp()
        {
            return _applicationProvider.NetApplication;
        }

        public int GetSelectedShapeCount()
        {
            var app = GetNetApp();
            if (app == null)
            {
                return 0;
            }

            try
            {
                // Use safe navigation for ActiveWindow as it might throw if no window is open
                var window = app.ActiveWindow;
                if (window == null)
                {
                    return 0;
                }

                var selection = window.Selection;
                if (selection == null)
                {
                    return 0;
                }

                if (selection.Type == NetOffice.PowerPointApi.Enums.PpSelectionType.ppSelectionShapes)
                {
                    var shapeRange = selection.ShapeRange;
                    if (shapeRange != null)
                    {
                        return shapeRange.Count;
                    }
                }

                return 0;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to get selected shape count: {ex.Message}");
                return 0;
            }
        }

        public NETOP.Selection GetSelection()
        {
            var app = GetNetApp();
            try
            {
                return app?.ActiveWindow?.Selection;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to get selection: {ex.Message}");
                return null;
            }
        }

        public bool HasShapesSelected()
        {
            return GetSelectedShapeCount() > 0;
        }
    }
}

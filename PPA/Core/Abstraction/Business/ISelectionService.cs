using NETOP = NetOffice.PowerPointApi;

namespace PPA.Core.Abstraction.Business
{
	/// <summary>
	/// Selection service interface for handling PowerPoint selection state
	/// </summary>
	public interface ISelectionService
	{
		/// <summary>
		/// Gets the count of currently selected shapes
		/// </summary>
		/// <returns> Number of selected shapes </returns>
		int GetSelectedShapeCount();

		/// <summary>
		/// Gets the current selection object
		/// </summary>
		/// <returns> Current selection or null </returns>
		NETOP.Selection GetSelection();

		/// <summary>
		/// Checks if the current selection contains shapes
		/// </summary>
		/// <returns> True if shapes are selected </returns>
		bool HasShapesSelected();
	}
}

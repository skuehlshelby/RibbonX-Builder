Imports System.Drawing
Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

	Public Interface IItemDimensions(Of T)

		Function WithItemDimensions(dimensions As Size) As T

		Function WithItemDimensions(height As Integer, width As Integer) As T

		Function WithItemDimensions(dimensions As Size, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As T

		Function WithItemDimensions(height As Integer, width As Integer, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As T

	End Interface

End Namespace
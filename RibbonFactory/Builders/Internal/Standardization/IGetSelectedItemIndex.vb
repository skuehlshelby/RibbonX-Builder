Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetSelectedItemIndex(Of TBuilder)

        Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As TBuilder

    End Interface

End Namespace
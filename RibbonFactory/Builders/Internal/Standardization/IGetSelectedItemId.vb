Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetSelectedItemId(Of T)

        Function GetSelectedItemIdFrom(callback As FromControl(Of String), setSelected As SelectionChanged) As T

    End Interface

End Namespace
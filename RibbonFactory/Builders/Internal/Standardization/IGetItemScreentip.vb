Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetItemScreentip(Of T)

        Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As T

    End Interface

End Namespace
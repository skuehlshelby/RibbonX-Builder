Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IGetItemSupertip(Of T)

        Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As T

    End Interface

End Namespace
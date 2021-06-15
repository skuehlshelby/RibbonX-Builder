Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemSuperTipSource(Of T As Builder)

        Function GetItemSuperTipFrom(superTipSource As FromControlAndIndex(Of String)) As T

    End Interface
End Namespace
Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetSelectedItemIdSource(Of T As Builder)

        Function GetSelectedItemIdFrom(itemIdSource As FromControlAndIndex(Of String)) As T

    End Interface
End Namespace
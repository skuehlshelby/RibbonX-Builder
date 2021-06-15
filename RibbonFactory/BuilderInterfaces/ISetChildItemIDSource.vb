Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemIdSource(Of T As Builder)

        Function GetItemIdFrom(idSource As FromControlAndIndex(Of String)) As T

    End Interface
End Namespace
Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemLabelSource(Of T As Builder)

        Function GetItemLabelFrom(labelSource As FromControlAndIndex(Of String)) As T

    End Interface
End Namespace
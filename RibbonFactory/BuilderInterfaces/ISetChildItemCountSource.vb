Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemCountSource(Of T As Builder)

        Function GetItemCountFrom(itemCountSource As FromControl(Of Integer)) As T

    End Interface
End Namespace
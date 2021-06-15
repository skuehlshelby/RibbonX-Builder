Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetSelectedItemIndexSource(Of T As Builder)

        Function GetSelectedItemIndexFrom(itemIndexSource As FromControl(Of Integer)) As T

    End Interface
End Namespace
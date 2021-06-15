Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetChildItemLabelVisibility(Of T As Builder)

        Function ShowLabelsOnChildItems() As T

        Function HideLabelsOnChildItems() As T

    End Interface
End Namespace
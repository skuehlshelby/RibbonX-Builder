Imports RibbonFactory.Builders

Namespace BuilderInterfaces

    Public Interface ISetChildItemImageVisibility(Of T As Builder)

        Function ShowImagesOnChildItems() As T

        Function HideImagesOnChildItems() As T

    End Interface

End Namespace
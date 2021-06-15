Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetImageVisibility(Of T As Builder)
        Function ShowImage() As T

        Function ShowImage(getShowImage As FromControl(Of Boolean)) As T

        Function HideImage() As T

        Function HideImage(getShowImage As FromControl(Of Boolean)) As T
    End Interface
End Namespace
Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetVisibility(Of T As Builder)
        Function Visible() As T

        Function Visible(callback As FromControl(Of Boolean)) As T

        Function Invisible() As T

        Function Invisible(callback As FromControl(Of Boolean)) As T
    End Interface
End NameSpace
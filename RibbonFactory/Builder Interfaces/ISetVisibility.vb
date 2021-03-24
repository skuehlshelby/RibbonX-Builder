Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetVisibility(Of T As RibbonElement)
        Function Visible() As Builder(Of T)

        Function Visible(callback As FromControl(Of Boolean)) As Builder(Of T)

        Function Invisible() As Builder(Of T)

        Function Invisible(callback As FromControl(Of Boolean)) As Builder(Of T)
    End Interface
End NameSpace
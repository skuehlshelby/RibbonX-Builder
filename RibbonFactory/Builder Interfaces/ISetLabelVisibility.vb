Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetLabelVisibility(Of T As RibbonElement)
        Function ShowLabel() As Builder(Of T)

        Function HideLabel() As Builder(Of T)
    End Interface
End NameSpace
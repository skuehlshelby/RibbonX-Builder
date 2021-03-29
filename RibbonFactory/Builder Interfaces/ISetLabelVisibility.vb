Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetLabelVisibility(Of T As Builder)
        Function ShowLabel() As T

        Function HideLabel() As T
    End Interface
End NameSpace
Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetLabelVisibility(Of T As Builder)
        Function ShowLabel() As T

        Function ShowLabel(callback As FromControl(Of Boolean)) As T

        Function HideLabel() As T
        
        Function HideLabel(callback As FromControl(Of Boolean)) As T
    End Interface
End NameSpace
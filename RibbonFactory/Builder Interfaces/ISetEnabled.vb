Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetEnabled(Of T As Builder)
        Function Enabled() As T
        Function Enabled(callback As FromControl(Of Boolean)) As T
        Function Disabled() As T
        Function Disabled(callback As FromControl(Of Boolean)) As T
    End Interface
End NameSpace
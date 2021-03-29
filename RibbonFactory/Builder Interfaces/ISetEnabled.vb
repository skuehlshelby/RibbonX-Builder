Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetEnabled(Of T As Builder)
        Function Enabled() As T
        Function Disabled() As T
    End Interface
End NameSpace
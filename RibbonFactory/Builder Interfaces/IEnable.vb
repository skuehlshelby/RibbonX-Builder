Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface IEnable(Of T As RibbonElement)
        Function Enabled() As Builder(Of T)
        Function Disabled() As Builder(Of T)
    End Interface
End NameSpace
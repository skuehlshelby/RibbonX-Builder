Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ILabel(Of T As RibbonElement) 
        Function WithLabel(label As String) As Builder(Of T)
    End Interface
End NameSpace
Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetLabel(Of T As Builder)
        Function WithLabel(label As String) As T
    End Interface
End NameSpace
Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetLabel(Of T As Builder)
        Function WithLabel(label As String) As T
    End Interface
End NameSpace
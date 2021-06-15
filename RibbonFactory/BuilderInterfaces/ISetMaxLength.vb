Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetMaxLength(Of T As Builder)
        Function OfWidth(limit As Byte) As T
    End Interface
End NameSpace
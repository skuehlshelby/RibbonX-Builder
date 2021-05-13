Imports RibbonFactory.Builders

Namespace BuilderInterfaces
    Public Interface ISetCharacterLimit(Of T As Builder)
        Function WithCharacterLimit(limit As Byte) As T
    End Interface
End NameSpace
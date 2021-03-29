Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetCharacterLimit(Of T As Builder)
        Function WithCharacterLimit(limit As Byte) As T
    End Interface
End NameSpace
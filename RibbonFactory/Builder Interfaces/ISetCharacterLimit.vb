Imports RibbonFactory.Builders

Namespace Builder_Interfaces
    Public Interface ISetCharacterLimit(Of T As RibbonElement)
        Function WithCharacterLimit(limit As Byte) As Builder(Of T)
    End Interface
End NameSpace
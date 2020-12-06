Public Interface IEditable
    Property Text As String
    ReadOnly Property CharacterLimit As Byte
End Interface
Public Interface IEditable(Of Out T)
    Function WithCharacterLimit(ByVal Limit As Byte) As T
End Interface

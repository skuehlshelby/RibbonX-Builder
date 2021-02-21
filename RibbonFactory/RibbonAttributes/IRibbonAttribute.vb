Namespace RibbonAttributes
    Public Interface IRibbonAttribute
        ReadOnly Property XML As String
    End Interface
    Public Interface IRibbonAttribute(Of T)
        Inherits IRibbonAttribute
        Function GetValue() As T
    End Interface
End NameSpace
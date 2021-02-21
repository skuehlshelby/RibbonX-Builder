Namespace RibbonAttributes
    Public Interface IMutableRibbonAttribute(Of T)
        Inherits IRibbonAttribute(Of T)
        Sub SetValue(value As T)
    End Interface
End NameSpace
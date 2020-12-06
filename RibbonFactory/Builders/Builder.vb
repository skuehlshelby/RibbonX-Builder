Public MustInherit Class Builder(Of T As RibbonElement)
    Implements IBuilder(Of T)
    Public MustOverride Function Build() As T Implements IBuilder(Of T).Build
End Class

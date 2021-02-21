Namespace Component_Interfaces
    Public Interface IContainer
        ReadOnly Property Items As List(Of RibbonElement)
    End Interface
    Public Interface IContainer(Of In T)
        Inherits IContainer
        Sub Add(ParamArray Items() As T)
    End Interface
End NameSpace
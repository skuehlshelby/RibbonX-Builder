Imports RibbonFactory.Builders
Imports RibbonFactory.Enums

Namespace Builder_Interfaces
    Public Interface ISetOrientation(Of T As Builder)
        Function Horizontal() As T
        Function Vertical() As T
    End Interface
End NameSpace
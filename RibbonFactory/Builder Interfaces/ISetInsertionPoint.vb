Imports RibbonFactory.Builders
Imports RibbonFactory.Enums.MSO

Namespace Builder_Interfaces

    Public Interface ISetInsertionPoint(Of T As Builder)
        Function InsertBefore(builtInControl As MSO) As T
        Function InsertBefore(qualifiedControl As RibbonElement) As T
        Function InsertAfter(builtInControl As MSO) As T
        Function InsertAfter(qualifiedControl As RibbonElement) As T
    End Interface
End NameSpace
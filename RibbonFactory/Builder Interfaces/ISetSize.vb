Imports RibbonFactory.Builders
Imports RibbonFactory.Enums

Namespace Builder_Interfaces
    Public Interface ISetSize(Of T As Builder)
        Overloads Function Large() As T
        Overloads Function Large(callback As FromControl(Of ControlSize)) As T
        Overloads Function Normal() As T
        Overloads Function Normal(callback As FromControl(Of ControlSize)) As T
    End Interface
End NameSpace
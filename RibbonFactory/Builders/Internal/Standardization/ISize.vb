Imports RibbonX.Callbacks
Imports RibbonX.SimpleTypes

Namespace Builders.Internal.Standardization

    Public Interface ISize(Of T)
        Overloads Function Large() As T
        Overloads Function Large(callback As FromControl(Of ControlSize)) As T
        Overloads Function Normal() As T
        Overloads Function Normal(callback As FromControl(Of ControlSize)) As T

    End Interface

End Namespace
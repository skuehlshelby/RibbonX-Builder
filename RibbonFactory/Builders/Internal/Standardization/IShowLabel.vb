Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IShowLabel(Of T)

        Function ShowLabel() As T

        Function ShowLabel(callback As FromControl(Of Boolean)) As T

        Function HideLabel() As T

        Function HideLabel(callback As FromControl(Of Boolean)) As T

    End Interface

End Namespace
Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IVisible(Of T)

        Function Visible() As T

        Function Visible(callback As FromControl(Of Boolean)) As T

        Function Invisible() As T

        Function Invisible(callback As FromControl(Of Boolean)) As T

    End Interface

End Namespace
Imports RibbonX.Callbacks

Namespace Builders.Internal.Standardization

    Public Interface IShowImage(Of T)

        Function ShowImage() As T

        Function ShowImage(getShowImage As FromControl(Of Boolean)) As T

        Function HideImage() As T

        Function HideImage(getShowImage As FromControl(Of Boolean)) As T

    End Interface

End Namespace
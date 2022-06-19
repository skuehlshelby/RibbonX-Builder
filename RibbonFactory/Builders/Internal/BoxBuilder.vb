Imports RibbonX.Builders.Internal.Base
Imports RibbonX.Callbacks
Imports RibbonX.Controls
Imports RibbonX.Controls.Base

Namespace Builders

    Friend NotInheritable Class BoxBuilder
        Inherits BuilderBase(Of Box)
        Implements IBoxBuilder

        Public Sub New(Optional template As RibbonElement = Nothing)
            MyBase.New(template)
        End Sub

        Public Function Horizontal() As IBoxBuilder Implements IBoxBuilder.Horizontal
            HorizontalBase()
            Return Me
        End Function

        Public Function Vertical() As IBoxBuilder Implements IBoxBuilder.Vertical
            VerticalBase()
            Return Me
        End Function

        Public Function Visible() As IBoxBuilder Implements IBoxBuilder.Visible
            VisibleBase()
            Return Me
        End Function

        Public Function Visible(callback As FromControl(Of Boolean)) As IBoxBuilder Implements IBoxBuilder.Visible
            VisibleBase(callback)
            Return Me
        End Function

        Public Function Invisible() As IBoxBuilder Implements IBoxBuilder.Invisible
            InvisibleBase()
            Return Me
        End Function

        Public Function Invisible(callback As FromControl(Of Boolean)) As IBoxBuilder Implements IBoxBuilder.Invisible
            InvisibleBase(callback)
            Return Me
        End Function

    End Class

End Namespace
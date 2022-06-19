Imports RibbonX.Callbacks
Imports RibbonX.Controls.Base

Namespace Builders.Internal.Standardization

    Public Interface IOnActionToggle(Of TBuilder, TRibbonElement As RibbonElement)

        Function ThatDoes(action As Action(Of TRibbonElement), callback As ButtonPressed) As TBuilder

    End Interface

End Namespace
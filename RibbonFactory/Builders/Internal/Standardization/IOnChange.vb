Imports RibbonX.Callbacks
Imports RibbonX.Controls.Base

Namespace Builders.Internal.Standardization

    Public Interface IOnChange(Of TRibbonElement As RibbonElement, Out TBuilder)

        Function ThatDoes(action As Action(Of TRibbonElement), callback As TextChanged) As TBuilder

    End Interface

End Namespace
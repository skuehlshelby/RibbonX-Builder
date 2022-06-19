Imports RibbonX.Controls
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Builders.Internal.Standardization
Imports RibbonX.Callbacks

Namespace Builders

    Public Interface IButtonBuilder
        Inherits IID(Of IButtonBuilder)
        Inherits IEnabled(Of IButtonBuilder)
        Inherits IVisible(Of IButtonBuilder)
        Inherits ILabelScreenTipSuperTip(Of IButtonBuilder)
        Inherits IShowLabel(Of IButtonBuilder)
        Inherits IShowImage(Of IButtonBuilder)
        Inherits ISize(Of IButtonBuilder)
        Inherits IImage(Of IButtonBuilder)
        Inherits IDescription(Of IButtonBuilder)
        Inherits IKeyTip(Of IButtonBuilder)
        Inherits IInsert(Of IButtonBuilder)

        Function RouteClickTo(callBack As OnAction) As IButtonBuilder
        Function BeforeClick(ParamArray actions() As Action(Of Button, CancelableEventArgs)) As IButtonBuilder
        Function OnClick(ParamArray actions() As Action(Of Button)) As IButtonBuilder

    End Interface

End Namespace
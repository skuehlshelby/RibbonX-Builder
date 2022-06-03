Imports RibbonX.Controls.Events

Namespace BuilderInterfaces.API

    Public Interface IToggleButtonBuilder
        Inherits IID(Of IToggleButtonBuilder)
        Inherits IVisible(Of IToggleButtonBuilder)
        Inherits IEnabled(Of IToggleButtonBuilder)
        Inherits IInsert(Of IToggleButtonBuilder)
        Inherits ILabelScreenTipSuperTip(Of IToggleButtonBuilder)
        Inherits IShowLabel(Of IToggleButtonBuilder)
        Inherits IKeyTip(Of IToggleButtonBuilder)
        Inherits IImage(Of IToggleButtonBuilder)
        Inherits IShowImage(Of IToggleButtonBuilder)
        Inherits IDescription(Of IToggleButtonBuilder)
        Inherits ISize(Of IToggleButtonBuilder)

        Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As IToggleButtonBuilder
        Function Unchecked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As IToggleButtonBuilder
        Function BeforeStateChange(ParamArray actions() As Action(Of BeforeToggleChangeEventArgs)) As IToggleButtonBuilder
        Function AfterStateChange(ParamArray actions() As Action(Of ToggleChangeEventArgs)) As IToggleButtonBuilder
    End Interface

End Namespace

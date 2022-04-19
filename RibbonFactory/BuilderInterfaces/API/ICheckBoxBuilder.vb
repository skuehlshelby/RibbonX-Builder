Imports RibbonFactory.Controls.Events

Namespace BuilderInterfaces.API

    Public Interface ICheckBoxBuilder
        Inherits IID(Of ICheckBoxBuilder)
        Inherits IEnabled(Of ICheckBoxBuilder)
        Inherits IVisible(Of ICheckBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of ICheckBoxBuilder)
        Inherits IInsert(Of ICheckBoxBuilder)
        Inherits IKeyTip(Of ICheckBoxBuilder)
        Inherits IDescription(Of ICheckBoxBuilder)

        Function Checked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As ICheckBoxBuilder
        Function Unchecked(getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As ICheckBoxBuilder

        Function BeforeStateChange(ParamArray actions() As Action(Of BeforeToggleChangeEventArgs)) As ICheckBoxBuilder
        Function AfterCheckStateChanged(ParamArray actions() As Action(Of ToggleChangeEventArgs)) As ICheckBoxBuilder
    End Interface

End Namespace

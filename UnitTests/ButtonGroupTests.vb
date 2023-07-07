Imports RibbonX
Imports RibbonX.Controls.BuiltIn
Imports Rx = RibbonX.RxApi

<TestClass()>
Public Class ButtonGroupTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim buttonGroup As IButtonGroup = Rx.ButtonGroup(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim buttonGroup As IButtonGroup = Rx.ButtonGroup()
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyButtonGroup(ButtonTests.BuildButton(),
                                                         SeparatorTests.BuildReadonlySeparator(),
                                                         ToggleButtonTests.BuildToggleButtonII()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(Rx.ButtonGroup())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim buttonGroup As IButtonGroup = BuildButtonGroup(ButtonTests.BuildButton(),
                                                         SeparatorTests.BuildReadonlySeparator(),
                                                         ToggleButtonTests.BuildToggleButtonII())

        Assert.AreEqual(buttonGroup.Visible, True)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyButtonGroupII(ButtonTests.BuildButton(),
                                                         SeparatorTests.BuildReadonlySeparator(),
                                                         ToggleButtonTests.BuildToggleButtonII()))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IButtonGroup = BuildButtonGroup(ButtonTests.BuildButton(),
                                                         SeparatorTests.BuildReadonlySeparator(),
                                                         ToggleButtonTests.BuildToggleButtonII())

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IButtonGroup = BuildButtonGroupII(ButtonTests.BuildButton(),
                                                         SeparatorTests.BuildReadonlySeparator(),
                                                         ToggleButtonTests.BuildToggleButtonII())

        Assert.That.SharedPropertiesAreEqual(control, Rx.Gallery(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildReadonlyButtonGroup(ParamArray controls() As IButtonGroupAddable) As IButtonGroup
        Return Rx.ButtonGroup(Sub(b) b.InsertAfter(Excel.About).Visible().WithControls(controls))
    End Function

    Public Shared Function BuildReadonlyButtonGroupII(ParamArray controls() As IButtonGroupAddable) As IButtonGroup
        Return Rx.ButtonGroup(Sub(b) b.InsertBefore(Excel.About).Invisible().WithControls(controls))
    End Function

    Public Shared Function BuildButtonGroup(ParamArray controls() As IButtonGroupAddable) As IButtonGroup
        Return Rx.ButtonGroup(Sub(b) b.InsertAfter(Excel.About).Visible(AddressOf GetVisibleShared).WithControls(controls))
    End Function

    Public Shared Function BuildButtonGroupII(ParamArray controls() As IButtonGroupAddable) As IButtonGroup
        Return Rx.ButtonGroup(Sub(b) b.InsertBefore(Excel.About).Invisible(AddressOf GetVisibleShared).WithControls(controls))
    End Function
End Class
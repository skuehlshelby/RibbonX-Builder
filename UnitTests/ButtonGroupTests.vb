Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn

<TestClass()>
Public Class ButtonGroupTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyButtonGroup(ButtonGroupControls.
                                                          From(ButtonTests.BuildButton()).
                                                          Add(SeparatorTests.BuildReadonlySeparator()).
                                                          Add(ToggleButtonTests.BuildToggleButtonII())))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New ButtonGroup())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim buttonGroup As ButtonGroup = BuildButtonGroup(ButtonGroupControls.
                                                          From(ButtonTests.BuildButton()).
                                                          Add(SeparatorTests.BuildReadonlySeparator()).
                                                          Add(ToggleButtonTests.BuildToggleButtonII()))

        Assert.AreEqual(buttonGroup.Visible, True)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyButtonGroupII(ButtonGroupControls.
                                                          From(ButtonTests.BuildButton()).
                                                          Add(SeparatorTests.BuildReadonlySeparator()).
                                                          Add(ToggleButtonTests.BuildToggleButtonII())))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As ButtonGroup = BuildButtonGroup(ButtonGroupControls.
                                                          From(ButtonTests.BuildButton()).
                                                          Add(SeparatorTests.BuildReadonlySeparator()).
                                                          Add(ToggleButtonTests.BuildToggleButtonII()))

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As ButtonGroup = BuildButtonGroupII(ButtonGroupControls.
                                                          From(ButtonTests.BuildButton()).
                                                          Add(SeparatorTests.BuildReadonlySeparator()).
                                                          Add(ToggleButtonTests.BuildToggleButtonII()))

        Assert.That.SharedPropertiesAreEqual(control, New Gallery(template:=control))
    End Sub

    Public Shared Function BuildReadonlyButtonGroup(controls As ButtonGroupControls) As ButtonGroup
        Return New ButtonGroup(config:=Sub(b) b.InsertAfter(Excel.About).Visible(), items:=controls)
    End Function

    Public Shared Function BuildReadonlyButtonGroupII(controls As ButtonGroupControls) As ButtonGroup
        Return New ButtonGroup(config:=Sub(b) b.InsertBefore(Excel.About).Invisible(), items:=controls)
    End Function

    Public Shared Function BuildButtonGroup(controls As ButtonGroupControls) As ButtonGroup
        Return New ButtonGroup(config:=Sub(b) b.InsertAfter(Excel.About).Visible(AddressOf GetVisibleShared), items:=controls)
    End Function

    Public Shared Function BuildButtonGroupII(controls As ButtonGroupControls) As ButtonGroup
        Return New ButtonGroup(config:=Sub(b) b.InsertBefore(Excel.About).Invisible(AddressOf GetVisibleShared), items:=controls)
    End Function
End Class
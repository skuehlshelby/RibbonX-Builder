Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass>
Public Class SplitButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As SplitButton = New SplitButton(ButtonTests.BuildButton(), New Menu(), template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As SplitButton = New SplitButton(ButtonTests.BuildButtonII(), New Menu(), config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildSplitButton())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New SplitButton(New ToggleButton(), New Menu()))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim button As Button = ButtonTests.BuildButton()
        Dim otherButton As Button = ButtonTests.BuildButton()
        Dim items As MenuControls = MenuControls.From(otherButton)
        Dim menu As Menu = New Menu(config:=Sub(mb) mb.WithLargeItems(), items:=items, template:=button)

        Dim control As SplitButton = New SplitButton(button, menu, template:=button, config:=Sub(sb) sb.
                                                                           Visible(AddressOf GetVisibleShared).
                                                                           Enabled(AddressOf GetEnabledShared).
                                                                           Large(AddressOf GetSizeShared).
                                                                           ShowLabel(AddressOf GetShowLabelShared).
                                                                           WithKeyTip("J1", AddressOf GetKeyTipShared))

        Assert.IsTrue(control.Enabled)
        Assert.IsTrue(control.Visible)
        Assert.IsTrue(control.ShowLabel)
        Assert.AreEqual(control.Size, ControlSize.Large)
        Assert.AreEqual(control.KeyTip, CType("J1", KeyTip))
        Assert.AreEqual(control.Button, button)
        Assert.AreEqual(control.Menu, menu)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlySplitButton())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As SplitButton = BuildSplitButtonII()

        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.KeyTip, CType("A", KeyTip))
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Size, ControlSize.Large)

        Debug.WriteLine(control)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As SplitButton = BuildReadonlySplitButtonII()

        Assert.That.SharedPropertiesAreEqual(control, New EditBox(template:=control))
    End Sub

    Public Shared Function BuildReadonlySplitButton() As SplitButton
        Dim button As Button = ButtonTests.BuildReadonlyButton()
        Dim otherButton As Button = ButtonTests.BuildReadonlyButton()
        Dim menu As Menu = MenuTests.BuildMenu(MenuControls.From(otherButton))

        Return New SplitButton(button, menu, template:=button, config:=Sub(sb) sb.
                                                                           InsertAfter(Excel.AlignLeftToRightMenu).
                                                                           Visible().
                                                                           Enabled().
                                                                           Large().
                                                                           ShowLabel().
                                                                           WithKeyTip("J1"))
    End Function

    Public Shared Function BuildReadonlySplitButtonII() As SplitButton
        Dim button As ToggleButton = ToggleButtonTests.BuildToggleButton()
        Dim otherButton As Button = ButtonTests.BuildReadonlyButtonII()
        Dim menu As Menu = MenuTests.BuildMenu(MenuControls.From(otherButton))

        Return New SplitButton(button, menu, template:=button, config:=Sub(sb) sb.
                                                                           InsertBefore(Excel.AlignLeftToRightMenu).
                                                                           Invisible().
                                                                           Disabled().
                                                                           Normal().
                                                                           HideLabel())
    End Function

    Public Shared Function BuildSplitButton() As SplitButton
        Dim button As Button = ButtonTests.BuildButton()
        Dim otherButton As Button = ButtonTests.BuildButton()
        Dim menu As Menu = MenuTests.BuildMenu(MenuControls.From(otherButton))

        Return New SplitButton(button, menu, template:=button, config:=Sub(sb) sb.
                                                                           InsertAfter(Excel.AlignLeftToRightMenu).
                                                                           Visible(AddressOf GetVisibleShared).
                                                                           Enabled(AddressOf GetEnabledShared).
                                                                           Large(AddressOf GetSizeShared).
                                                                           ShowLabel(AddressOf GetShowLabelShared).
                                                                           WithKeyTip("J1", AddressOf GetKeyTipShared))
    End Function

    Public Shared Function BuildSplitButtonII() As SplitButton
        Dim button As ToggleButton = ToggleButtonTests.BuildToggleButtonII()
        Dim otherButton As Button = ButtonTests.BuildButtonII()
        Dim menu As Menu = MenuTests.BuildMenu(MenuControls.From(otherButton))

        Return New SplitButton(button, menu, template:=button, config:=Sub(sb) sb.
                                                                           InsertAfter(Excel.AlignLeftToRightMenu).
                                                                           Invisible(AddressOf GetVisibleShared).
                                                                           Disabled(AddressOf GetEnabledShared).
                                                                           Normal(AddressOf GetSizeShared).
                                                                           HideLabel(AddressOf GetShowLabelShared))
    End Function

End Class

Imports RibbonX
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass>
Public Class SplitButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = RxApi.Menu(Sub(mb) mb.WithLargeItems().FromTemplate(button).WithControls(otherButton))
        Dim control As ISplitButton = RxApi.SplitButton(Sub(b) b.WithButtonAndMenu(button, menu).FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = RxApi.Menu(Sub(mb) mb.WithLargeItems().FromTemplate(button).WithControls(otherButton))
        Dim control As ISplitButton = RxApi.SplitButton(Sub(b) b.WithButtonAndMenu(button, menu))
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = RxApi.Menu(Sub(mb) mb.WithLargeItems().FromTemplate(button).WithControls(otherButton))
        Assert.That.ValidRibbonXIsProduced(RxApi.SplitButton(Sub(b) b.WithButtonAndMenu(button, menu)))
        Assert.That.ValidRibbonXIsProduced(BuildSplitButton())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = RxApi.Menu(Sub(mb) mb.WithLargeItems().FromTemplate(button).WithControls(otherButton))
        Assert.That.NoPropertiesAreNull(RxApi.SplitButton(Sub(b) b.WithButtonAndMenu(button, menu)))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = RxApi.Menu(Sub(mb) mb.WithLargeItems().FromTemplate(button).WithControls(otherButton))

        Dim control As ISplitButton = RxApi.SplitButton(Sub(b) b.
                                                            WithButtonAndMenu(button, menu).
                                                            FromTemplate(button).
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
        Dim control As ISplitButton = BuildSplitButtonII()

        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.KeyTip, CType("A", KeyTip))
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Size, ControlSize.Large)

        Debug.WriteLine(control)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As ISplitButton = BuildReadonlySplitButtonII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.EditBox(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildReadonlySplitButton() As ISplitButton
        Dim button As IButton = ButtonTests.BuildReadonlyButton()
        Dim otherButton As IButton = ButtonTests.BuildReadonlyButton()
        Dim menu As IMenu = MenuTests.BuildMenu(otherButton)

        Return RxApi.SplitButton(Sub(b) b.
                                     WithButtonAndMenu(button, menu).
                                     FromTemplate(button).
                                     InsertAfter(Excel.AlignLeftToRightMenu).
                                     Visible().
                                     Enabled().
                                     Large().
                                     ShowLabel().
                                     WithKeyTip("J1"))
    End Function

    Public Shared Function BuildReadonlySplitButtonII() As ISplitButton
        Dim button As IToggleButton = ToggleButtonTests.BuildToggleButton()
        Dim otherButton As IButton = ButtonTests.BuildReadonlyButtonII()
        Dim menu As IMenu = MenuTests.BuildMenu(otherButton)

        Return RxApi.SplitButton(Sub(b) b.
                                     WithButtonAndMenu(button, menu).
                                     FromTemplate(button).
                                     InsertBefore(Excel.AlignLeftToRightMenu).
                                     Invisible().
                                     Disabled().
                                     Normal().
                                     HideLabel())
    End Function

    Public Shared Function BuildSplitButton() As ISplitButton
        Dim button As IButton = ButtonTests.BuildButton()
        Dim otherButton As IButton = ButtonTests.BuildButton()
        Dim menu As IMenu = MenuTests.BuildMenu(otherButton)

        Return RxApi.SplitButton(Sub(b) b.
                                     WithButtonAndMenu(button, menu).
                                     FromTemplate(button).
                                    InsertAfter(Excel.AlignLeftToRightMenu).
                                    Visible(AddressOf GetVisibleShared).
                                    Enabled(AddressOf GetEnabledShared).
                                    Large(AddressOf GetSizeShared).
                                    ShowLabel(AddressOf GetShowLabelShared).
                                    WithKeyTip("J1", AddressOf GetKeyTipShared))
    End Function

    Public Shared Function BuildSplitButtonII() As ISplitButton
        Dim button As IToggleButton = ToggleButtonTests.BuildToggleButtonII()
        Dim otherButton As IButton = ButtonTests.BuildButtonII()
        Dim menu As IMenu = MenuTests.BuildMenu(otherButton)

        Return RxApi.SplitButton(Sub(b) b.
                                     WithButtonAndMenu(button, menu).
                                     FromTemplate(button).
                                    InsertAfter(Excel.AlignLeftToRightMenu).
                                    Invisible(AddressOf GetVisibleShared).
                                    Disabled(AddressOf GetEnabledShared).
                                    Normal(AddressOf GetSizeShared).
                                    HideLabel(AddressOf GetShowLabelShared))
    End Function

End Class

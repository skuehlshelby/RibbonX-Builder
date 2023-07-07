Imports System.Drawing
Imports RibbonX
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass>
Public Class ToggleButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As IToggleButton = RxApi.ToggleButton(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As IToggleButton = RxApi.ToggleButton(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyToggleButton())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.ToggleButton())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IToggleButton =
            RxApi.ToggleButton(Sub(builder) builder.
                                Disabled().
                                Invisible().
                                Normal().
                                WithLabel("Label").
                                WithScreenTip("ScreenTip").
                                WithSuperTip("SuperTip").
                                WithDescription("Description").
                                InsertBefore(Excel.About).
                                WithKeyTip("P").
                                WithImage(Number.Three).
                                HideImage().
                                HideLabel())

        Assert.IsFalse(control.Enabled)
        Assert.IsFalse(control.Visible)
        Assert.IsFalse(control.ShowLabel)
        Assert.IsFalse(control.ShowImage)
        Assert.AreEqual(control.Size, ControlSize.Normal)
        Assert.AreEqual(control.Label, "Label")
        Assert.AreEqual(control.ScreenTip, "ScreenTip")
        Assert.AreEqual(control.SuperTip, "SuperTip")
        Assert.AreEqual(control.Description, "Description")
        Assert.AreEqual(control.Image, Number.Three)
        Assert.AreEqual(control.KeyTip, CType("P", KeyTip))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyToggleButtonII())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IToggleButton = BuildToggleButton()

        Assert.That.PropertyMayBeModified(control, Function(c) control.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(c) control.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(c) control.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(c) control.ShowImage, Not control.ShowImage)
        Assert.That.PropertyMayBeModified(control, Function(c) control.Size, ControlSize.Normal)
        Assert.That.PropertyMayBeModified(control, Function(c) control.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) control.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) control.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) control.Description, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) control.KeyTip, CType("K", KeyTip))
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IToggleButton = BuildToggleButtonII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.Button(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildReadonlyToggleButton() As IToggleButton
        Return RxApi.ToggleButton(Sub(builder) builder.
                                Enabled().
                                Visible().
                                Large().
                                WithLabel("Label").
                                WithScreenTip("ScreenTip").
                                WithSuperTip("SuperTip").
                                WithDescription("Description").
                                InsertBefore(Excel.About).
                                WithKeyTip("P").
                                WithImage(Number.Three).
                                ShowImage().
                                ShowLabel())
    End Function

    Public Shared Function BuildReadonlyToggleButtonII() As IToggleButton
        Return RxApi.ToggleButton(Sub(builder) builder.
                        Disabled().
                        Invisible().
                        Normal().
                        WithLabel("Label", False).
                        WithScreenTip("ScreenTip").
                        WithSuperTip("SuperTip").
                        WithDescription("Description").
                        InsertAfter(Excel.About).
                        WithKeyTip("P").
                        WithImage(Number.Three).
                        HideImage().
                        HideLabel())
    End Function

    Public Shared Function BuildToggleButton() As IToggleButton
        Return RxApi.ToggleButton(Sub(builder) builder.
                        Enabled(AddressOf GetEnabledShared).
                        Visible(AddressOf GetVisibleShared).
                        Large(AddressOf GetSizeShared).
                        WithLabel("Label", AddressOf GetLabelShared).
                        WithScreenTip("ScreenTip", AddressOf GetScreenTipShared).
                        WithSuperTip("SuperTip", AddressOf GetSuperTipShared).
                        WithDescription("Description", AddressOf GetDescriptionShared).
                        WithKeyTip("P", AddressOf GetKeyTipShared).
                        WithImage(New Bitmap(1, 1), AddressOf GetImageShared).
                        ShowImage(AddressOf GetShowImageShared).
                        ShowLabel(AddressOf GetShowLabelShared))
    End Function

    Public Shared Function BuildToggleButtonII() As IToggleButton
        Return RxApi.ToggleButton(Sub(builder) builder.
                        Disabled(AddressOf GetEnabledShared).
                        Invisible(AddressOf GetVisibleShared).
                        Normal(AddressOf GetSizeShared).
                        WithLabel("Label", AddressOf GetLabelShared, False).
                        WithScreenTip("ScreenTip", AddressOf GetScreenTipShared).
                        WithSuperTip("SuperTip", AddressOf GetSuperTipShared).
                        WithDescription("Description", AddressOf GetDescriptionShared).
                        WithKeyTip("P", AddressOf GetKeyTipShared).
                        WithImage(New Bitmap(1, 1), AddressOf GetImageShared).
                        HideImage(AddressOf GetShowImageShared).
                        HideLabel(AddressOf GetShowLabelShared))
    End Function
End Class

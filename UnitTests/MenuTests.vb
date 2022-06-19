Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class MenuTests
    Inherits TestBase

    Private Const LABEL As String = "My Menu"
    Private Const SCREENTIP As String = "My Menu Screentip"
    Private Const SUPERTIP As String = "More Info About My Menu"
    Private Const DESCRIPTION As String = "Even More Info About My Menu"
    Private Const KEYTIP As String = "B"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As Menu = New Menu(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As Menu = New Menu(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildMenu(MenuControls.From(ButtonTests.BuildButton(), ButtonTests.BuildButtonII()).Add(ToggleButtonTests.BuildToggleButton())))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Menu())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As Menu = BuildReadonlyMenu(MenuControls.From(ButtonTests.BuildButton(), ButtonTests.BuildButtonII()))

        Assert.IsTrue(control.Enabled)
        Assert.IsTrue(control.Visible)
        Assert.IsTrue(control.ShowImage)
        Assert.IsTrue(control.ShowLabel)
        Assert.AreEqual(control.Label, LABEL)
        Assert.AreEqual(control.ScreenTip, SCREENTIP)
        Assert.AreEqual(control.SuperTip, SUPERTIP)
        Assert.AreEqual(control.Description, DESCRIPTION)
        Assert.AreEqual(control.KeyTip, KEYTIP)
        Assert.AreEqual(control.Size, ControlSize.Large)
        Assert.AreEqual(control.Image, "img/blah.jpg")
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyMenu(MenuControls.From(ButtonTests.BuildButton(), ButtonTests.BuildButtonII())))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As Menu = BuildMenu(MenuControls.From(SplitButtonTests.BuildSplitButtonII()).Add(ButtonTests.BuildButtonII()))

        Assert.That.PropertyMayBeModified(control, Function(c) c.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ShowImage, Not control.ShowImage)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Description, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Size, ControlSize.Large)
        Assert.That.PropertyMayBeModified(control, Function(c) c.KeyTip, CType("A", KeyTip))

        Debug.WriteLine(control.XML)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim menu As Menu = BuildMenu(MenuControls.From(ButtonTests.BuildButton(), ButtonTests.BuildButtonII()))

        Assert.That.SharedPropertiesAreEqual(menu, New Gallery(template:=menu))
    End Sub

    Public Shared Function BuildReadonlyMenu(controls As MenuControls) As Menu
        Return New Menu(config:=Sub(b) b.
                        InsertAfter(Excel.Calculator).
                        Enabled().
                        Visible().
                        Large().
                        WithLargeItems().
                        WithLabel(LABEL).
                        WithScreenTip(SCREENTIP).
                        WithSuperTip(SUPERTIP).
                        WithDescription(DESCRIPTION).
                        WithImage("img/blah.jpg").
                        WithKeyTip(KEYTIP),
                        items:=controls)
    End Function

    Public Shared Function BuildMenu(controls As MenuControls) As Menu
        Return New Menu(config:=Sub(b) b.
                        InsertAfter(Excel.Calculator).
                        Disabled(AddressOf GetEnabledShared).
                        Invisible(AddressOf GetVisibleShared).
                        Normal(AddressOf GetSizeShared).
                        WithLargeItems().
                        WithLabel(LABEL, AddressOf GetLabelShared).
                        HideLabel(AddressOf GetShowLabelShared).
                        WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                        WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                        WithDescription(DESCRIPTION, AddressOf GetDescriptionShared).
                        WithImage(BlankBitmap, AddressOf GetImageShared).
                        HideImage(AddressOf GetShowImageShared).
                        WithKeyTip(KEYTIP, AddressOf GetKeyTipShared),
                        items:=controls)
    End Function
End Class
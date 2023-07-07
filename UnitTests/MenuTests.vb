Imports RibbonX
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
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
        Dim control As IMenu = RxApi.Menu(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As IMenu = RxApi.Menu(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildMenu(ButtonTests.BuildButton(), ButtonTests.BuildButtonII(), ToggleButtonTests.BuildToggleButton()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.Menu())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IMenu = BuildReadonlyMenu(ButtonTests.BuildButton(), ButtonTests.BuildButtonII())

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
        Assert.AreEqual(control.Image, Common.DollarSign)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyMenu(ButtonTests.BuildButton(), ButtonTests.BuildButtonII()))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IMenu = BuildMenu(SplitButtonTests.BuildSplitButtonII(), ButtonTests.BuildButtonII())

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

        Debug.WriteLine(control.ToXml())
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim menu As IMenu = BuildMenu(ButtonTests.BuildButton(), ButtonTests.BuildButtonII())

        Assert.That.SharedPropertiesAreEqual(menu, RxApi.Gallery(Sub(b) b.FromTemplate(menu)))
    End Sub

    Public Shared Function BuildReadonlyMenu(ParamArray controls() As IMenuAddable) As IMenu
        Return RxApi.Menu(Sub(b) b.
                        InsertAfter(Excel.Calculator).
                        Enabled().
                        Visible().
                        Large().
                        WithLargeItems().
                        WithLabel(LABEL).
                        WithScreenTip(SCREENTIP).
                        WithSuperTip(SUPERTIP).
                        WithDescription(DESCRIPTION).
                        WithImage(Common.DollarSign).
                        WithKeyTip(KEYTIP).
                        WithControls(controls))
    End Function

    Public Shared Function BuildMenu(ParamArray controls() As IMenuAddable) As IMenu
        Return RxApi.Menu(Sub(b) b.
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
                        WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                        WithControls(controls))
    End Function
End Class
Imports RibbonX.Controls
Imports RibbonX.Controls.Base
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class GroupTests
    Inherits TestBase

    Private Const LABEL As String = "My Group"
    Private Const SCREENTIP As String = "My Group Screentip"
    Private Const SUPERTIP As String = "More Info About My Group"
    Private Const KEYTIP As String = "C"

    <TestMethod()>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As Group = New Group(template:=Nothing)
    End Sub

    <TestMethod()>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As Group = New Group(template:=Nothing)
    End Sub

    <TestMethod()>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(New Group())
    End Sub

    <TestMethod()>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Group())
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As Group = BuildReadonlyGroup(SplitButtonTests.BuildReadonlySplitButton())

        Assert.IsTrue(control.Visible)
        Assert.AreEqual(control.Image, Common.SadFace)
        Assert.AreEqual(control.Label, LABEL)
        Assert.AreEqual(control.ScreenTip, SCREENTIP)
        Assert.AreEqual(control.SuperTip, SUPERTIP)
        Assert.AreEqual(control.KeyTip, CType(KEYTIP, KeyTip))
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyGroup(ButtonTests.BuildButton()))
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As Group = BuildGroup(SplitButtonTests.BuildSplitButtonII())

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Image, Nothing)
        Assert.That.PropertyMayBeModified(control, Function(c) c.KeyTip, CType("J", KeyTip))
    End Sub

    <TestMethod()>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As Group = BuildReadonlyGroup()

        Assert.That.SharedPropertiesAreEqual(control, New Button(template:=control))
    End Sub

    Public Shared Function BuildReadonlyGroup(ParamArray controls() As RibbonElement) As Group
        Return New [Group](config:=Sub(b) b.
                           WithImage(Common.SadFace).
                           Visible().
                           WithKeyTip(KEYTIP).
                           WithLabel(LABEL).
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           InsertAfter(Excel.About),
                           controls:=controls)
    End Function

    Public Shared Function BuildGroup(ParamArray controls() As RibbonElement) As Group
        Return New [Group](config:=Sub(b) b.
                           WithImage(BlankBitmap(), AddressOf GetImageShared).
                           Invisible(AddressOf GetVisibleShared).
                           WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                           WithLabel(LABEL, AddressOf GetLabelShared).
                           WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                           WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared),
                           controls:=controls)
    End Function
End Class

Imports RibbonX
Imports RibbonX.Controls
Imports RibbonX.Images.RibbonImage

<TestClass()>
Public Class ItemTests
    Inherits TestBase

    <TestMethod()>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As IItem = RxApi.Item(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod()>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As IItem = RxApi.Item(Nothing)
    End Sub

    <TestMethod()>
    Public Overrides Sub ProducesLegalRibbonX()
        'Produces no RibbonX
    End Sub

    <TestMethod()>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.Item())
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IItem = BuildItem()

        Assert.AreEqual("Item Label", control.Label)
        Assert.AreEqual("Item ScreenTip", control.ScreenTip)
        Assert.AreEqual("Item Supertip", control.SuperTip)
        Assert.AreEqual(control.Image.ImageType, RibbonImageType.IPictureDisp)
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        'Item properties can always be modified.
    End Sub

    <TestMethod()>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IItem = BuildItem()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Image, Nothing)
    End Sub

    <TestMethod()>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IItem = BuildItem()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.DropDown(Sub(b) b.FromTemplate(control)))
    End Sub

    Public Shared Function BuildItem() As IItem
        Return RxApi.Item(Sub(b) b.
                            WithLabel("Item Label", False).
                            WithScreenTip("Item ScreenTip").
                            WithSuperTip("Item Supertip").
                            WithImage(BlankBitmap()))
    End Function

    Public Shared Function BuildItemII() As IItem
        Return RxApi.Item(Sub(b) b.
                            WithLabel("Other Item Label", False).
                            WithScreenTip("Other Item ScreenTip").
                            WithSuperTip("Other Item Supertip").
                            WithImage(BlankBitmap()))
    End Function
End Class

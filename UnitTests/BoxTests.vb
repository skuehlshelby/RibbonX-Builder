Imports RibbonX
Imports RibbonX.SimpleTypes
Imports Rx = RibbonX.RxApi

<TestClass()>
Public Class BoxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim box As IBox = Rx.Box(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim box As IBox = RX.Box()
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(Rx.Instance.BoxHorizontal(ButtonTests.BuildButton(), IGalleryTests.BuildReadonlyGallery()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(Rx.Box())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim box As IBox = Rx.Box(Sub(b) b.Horizontal().Visible().WithControls(Rx.Button(), Rx.CheckBox(), Rx.EditBox()))

        Assert.AreEqual(box.Visible, True)
        Assert.AreEqual(box.BoxStyle, BoxStyle.Horizontal)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim box As IBox = Rx.Box(Sub(b) b.Horizontal().Visible())

        Assert.ThrowsException(Of PropertyNotSettableException)(Sub() box.Visible = False)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim box As IBox = Rx.Box(Sub(b) b.Horizontal().Visible(AddressOf GetVisible))

        box.Visible = False
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim box As IBox = Rx.Box(Sub(b) b.Vertical().Invisible())

        Assert.That.SharedPropertiesAreEqual(box, Rx.Button(Sub(b) b.FromTemplate(box)))
    End Sub
End Class
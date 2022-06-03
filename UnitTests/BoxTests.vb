Imports RibbonX.Enums
Imports RibbonX.BuilderInterfaces.API
Imports RibbonX.Controls

<TestClass()>
Public Class BoxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim box As Box = New Box(Nothing, Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim box As Box = New Box(Nothing, Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        CreateSimpleRibbon(Box.Horizontal(New Button()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        ContainsNoNullValuesByDefaultHelper(New Box(Nothing, Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim box As Box = Box.Horizontal()

        Assert.AreEqual(box.Visible, True)
        Assert.AreEqual(box.BoxStyle, BoxStyle.horizontal)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim box As Box = New Box(Sub(bb As IBoxBuilder) bb.Visible().Horizontal(), Nothing)

        Assert.ThrowsException(Of Exception)(Sub() box.Visible = False)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim box As Box = New Box(Sub(bb As IBoxBuilder) bb.Visible(AddressOf GetVisible).Horizontal(), Nothing)

        box.Visible = False
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim box As Box = New Box(Sub(bb As IBoxBuilder) bb.Invisible().Vertical(), Nothing)

        Dim otherBox As Box = New Box(Nothing, Nothing, box)

        Assert.AreEqual(box.Visible, False)
        Assert.AreEqual(box.BoxStyle, BoxStyle.vertical)
    End Sub
End Class
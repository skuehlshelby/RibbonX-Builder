Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class BoxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim box As Box = New Box(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim box As Box = New Box(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(Box.Horizontal(ButtonTests.BuildButton(), ToggleButtonTests.BuildReadonlyToggleButton()))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Box())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim box As Box = Box.Horizontal()

        Assert.AreEqual(box.Visible, True)
        Assert.AreEqual(box.BoxStyle, BoxStyle.Horizontal)
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

        Assert.That.SharedPropertiesAreEqual(box, New Button(template:=box))
    End Sub
End Class
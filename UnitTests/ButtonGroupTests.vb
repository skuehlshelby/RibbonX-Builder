Imports RibbonX.Builders
Imports RibbonX.Controls

<TestClass()>
Public Class ButtonGroupTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(New ButtonGroup(items:=ButtonGroupControls.From(New Button())))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New ButtonGroup())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(config:=Sub(b) b.Visible())

        Assert.AreEqual(buttonGroup.Visible, True)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(Sub(bgb As IButtonGroupBuilder) bgb.Invisible())

        Assert.ThrowsException(Of Exception)(Sub() buttonGroup.Visible = True)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim buttonGroup As ButtonGroup = New ButtonGroup(Sub(bgb As IButtonGroupBuilder) bgb.Invisible(AddressOf GetVisible))

        buttonGroup.Visible = True
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim template As Button = New Button(Sub(bb) bb.Invisible())
        Dim buttonGroup As ButtonGroup = New ButtonGroup(Nothing, Nothing, template)

        Assert.AreEqual(buttonGroup.Visible, template.Visible)
    End Sub

End Class
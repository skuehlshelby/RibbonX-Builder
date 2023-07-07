Imports RibbonX
Imports RibbonX.Controls.BuiltIn

<TestClass()>
Public Class EditBoxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As IEditBox = RxApi.EditBox(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As IEditBox = RxApi.EditBox(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(RxApi.EditBox(Sub(builder) builder.
                           Visible().
                           Enabled().
                           WithLabel("Label").
                           AsWideAs("This Big").
                           InsertBefore(Excel.About).
                           WithMaximumInputLength(40).
                           WithText("Text", AddressOf GetText, AddressOf OnChange)))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.EditBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IEditBox = RxApi.EditBox(Sub(builder) builder.
                           Invisible().
                           Disabled().
                           WithLabel("Label").
                           HideLabel().
                           AsWideAs("This Big").
                           InsertBefore(Excel.About).
                           WithMaximumInputLength(40).
                           WithText("Text", AddressOf GetText, AddressOf OnChange))

        Assert.AreEqual(False, control.Visible)
        Assert.AreEqual(False, control.Enabled)
        Assert.AreEqual(False, control.ShowLabel)
        Assert.AreEqual("Label", control.Label)
        Assert.AreEqual("Label", control.ScreenTip)
        Assert.AreEqual("Text", control.Text)
        Assert.AreEqual(CByte(40), control.MaxLength)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim control As IEditBox = RxApi.EditBox(Sub(builder) builder.
                           Invisible().
                           Disabled().
                           WithLabel("Label").
                           HideLabel().
                           AsWideAs("This Big").
                           InsertBefore(Excel.About).
                           WithMaximumInputLength(40).
                           WithText("Text", AddressOf GetText, AddressOf OnChange))

        Assert.That.PropertiesCannotBeModified(control)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IEditBox = RxApi.EditBox(Sub(builder) builder.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           WithLabel("Label", AddressOf GetLabel).
                           HideLabel(AddressOf GetShowLabel).
                           WithText("Text", AddressOf GetText, AddressOf OnChange))

        control.Visible = False
        control.Enabled = False
        control.ShowLabel = False
        control.Label = String.Empty
        control.ScreenTip = String.Empty
        control.Text = String.Empty
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IEditBox = RxApi.EditBox(Sub(builder) builder.
                           Invisible().
                           Disabled().
                           WithLabel("Label").
                           HideLabel().
                           AsWideAs("This Big").
                           InsertBefore(Excel.About).
                           WithMaximumInputLength(40).
                           WithText("Text", AddressOf GetText, AddressOf OnChange))

        Dim otherControl As IButton = RxApi.Button(Sub(builder) builder.FromTemplate(control))

        Assert.That.SharedPropertiesAreEqual(control, otherControl)
    End Sub
End Class
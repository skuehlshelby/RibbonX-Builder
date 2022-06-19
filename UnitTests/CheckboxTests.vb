Imports RibbonX.Controls
Imports RibbonX.Controls.EventArgs
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class CheckboxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim checkbox As CheckBox = New CheckBox(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim checkbox As CheckBox = New CheckBox(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(New CheckBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithDescription("The Description")))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New CheckBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.
                           Disabled().
                           Invisible().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithDescription("The Description").
                           WithKeyTip("K2"))

        Assert.IsFalse(checkbox.Enabled)
        Assert.IsFalse(checkbox.Visible)
        Assert.AreEqual(checkbox.Label, "The Label")
        Assert.AreEqual(checkbox.ScreenTip, checkbox.Label)
        Assert.AreEqual(checkbox.SuperTip, "The Supertip")
        Assert.AreEqual(checkbox.Description, "The Description")
        Assert.AreEqual(checkbox.KeyTip, "K2")

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.
                           Invisible().
                           Disabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithDescription("The Description").
                           WithKeyTip("K2"))

        Assert.That.PropertiesCannotBeModified(checkbox)

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           WithLabel("Button", AddressOf GetLabel).
                           WithSuperTip("Super", AddressOf GetSuperTip).
                           WithDescription("Description", AddressOf GetDescription).
                           WithKeyTip("K2", AddressOf GetKeyTip).
                           Checked(AddressOf GetPressed, AddressOf OnButtonToggle).
                           BeforeStateChange(Sub() Debug.WriteLine("Before Toggle Fired")).
                           AfterCheckStateChanged(Sub() Debug.WriteLine("On Toggle Fired")))

        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Visible, Not checkbox.Visible)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Enabled, Not checkbox.Enabled)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Checked, Not checkbox.Checked)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Label, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Description, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.KeyTip, CType("A", KeyTip))

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As CheckBox = New CheckBox(Sub(cbb) cbb.WithDescription("The Description").
                                                   WithLabel("The Label").
                                                   WithId("MyID"))

        Assert.That.SharedPropertiesAreEqual(control, New Button(template:=control))
    End Sub

    <TestMethod()>
    Public Sub EventCanFire()

        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.Unchecked(AddressOf GetPressed, AddressOf OnButtonToggle))

        Dim fired As Boolean = False

        AddHandler checkbox.StateChanged, Sub(sender, e) fired = True

        checkbox.Checked = True

        Assert.IsTrue(fired)

    End Sub

    <TestMethod()>
    Public Sub EventCanBeCancelled()

        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.Unchecked(AddressOf GetPressed, AddressOf OnButtonToggle))

        AddHandler checkbox.BeforeStateChange, Sub(sender, e) e.Cancel()

        checkbox.Checked = True

        Assert.IsFalse(checkbox.Checked)

    End Sub

    <TestMethod()>
    Public Sub HandlersCanBeAddedAndRemoved()

        Dim checkbox As CheckBox = New CheckBox()
        Dim beforeChange As EventHandler(Of BeforeToggleChangeEventArgs) = New EventHandler(Of BeforeToggleChangeEventArgs)(Sub(sender, e) Return)
        Dim onChange As EventHandler(Of ToggleChangeEventArgs) = New EventHandler(Of ToggleChangeEventArgs)(Sub(sender, e) Return)

        AddHandler checkbox.BeforeStateChange, beforeChange
        AddHandler checkbox.StateChanged, onChange

        RemoveHandler checkbox.BeforeStateChange, beforeChange
        RemoveHandler checkbox.StateChanged, onChange
    End Sub

End Class
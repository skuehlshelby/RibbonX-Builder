Imports RibbonFactory.Controls
Imports RibbonFactory.Controls.Events

<TestClass()>
Public Class CheckboxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim checkbox As CheckBox = New CheckBox(Sub(cbb) cbb.Visible(), Nothing)

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim checkbox As CheckBox = New CheckBox(Nothing, Nothing)

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Dim ribbon As Ribbon = CreateSimpleRibbon(
            New CheckBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithDescription("The Description")))

        Assert.IsTrue(ribbon.GetErrors().None)

        Debug.WriteLine(ribbon.RibbonX)
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Dim checkbox As CheckBox = New CheckBox()

        ContainsNoNullValuesByDefaultHelper(checkbox)

        Debug.WriteLine(checkbox)
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

        Assert.ThrowsException(Of Exception)(Sub() checkbox.Visible = True)
        Assert.ThrowsException(Of Exception)(Sub() checkbox.Enabled = True)
        Assert.ThrowsException(Of Exception)(Sub() checkbox.Label = "New Label")
        Assert.ThrowsException(Of Exception)(Sub() checkbox.ScreenTip = "New Screentip")
        Assert.ThrowsException(Of Exception)(Sub() checkbox.SuperTip = "New Supertip")
        Assert.ThrowsException(Of Exception)(Sub() checkbox.KeyTip = "K5")
        Assert.ThrowsException(Of Exception)(Sub() checkbox.Description = "New Description")

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

        Assert.IsTrue(checkbox.Visible)
        checkbox.Visible = False
        Assert.IsFalse(checkbox.Visible)

        Assert.IsTrue(checkbox.Enabled)
        checkbox.Enabled = False
        Assert.IsFalse(checkbox.Enabled)

        Assert.IsTrue(checkbox.Checked)
        checkbox.Checked = False
        Assert.IsFalse(checkbox.Checked)

        Assert.AreEqual(checkbox.Label, "Button")
        checkbox.Label = "New Label"
        Assert.AreEqual(checkbox.Label, "New Label")

        Assert.AreEqual(checkbox.ScreenTip, "Button")
        checkbox.ScreenTip = "New ScreenTip"
        Assert.AreEqual(checkbox.ScreenTip, "New ScreenTip")

        Assert.AreEqual(checkbox.SuperTip, "Super")
        checkbox.SuperTip = "New SuperTip"
        Assert.AreEqual(checkbox.SuperTip, "New SuperTip")

        Assert.AreEqual(checkbox.Description, "Description")
        checkbox.Description = "New Description"
        Assert.AreEqual(checkbox.Description, "New Description")

        Assert.AreEqual(checkbox.KeyTip, "K2")
        checkbox.KeyTip = "A"
        Assert.AreEqual(checkbox.KeyTip, "A")

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim firstCheckbox As CheckBox = New CheckBox(Sub(cbb) cbb.WithDescription("The Description").WithLabel("The Label").WithId("MyID"))
        Dim secondCheckbox As CheckBox = New CheckBox(Nothing, firstCheckbox)

        Assert.AreEqual(secondCheckbox.Description, "The Description")
        Assert.AreEqual(secondCheckbox.Label, "The Label")
        Assert.AreNotEqual(secondCheckbox.ID, firstCheckbox.ID)

        Debug.WriteLine(firstCheckbox)
        Debug.WriteLine(secondCheckbox)
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
﻿
Imports RibbonX
Imports RibbonX.SimpleTypes
Imports Rx = RibbonX.RxApi

<TestClass()>
Public Class CheckboxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim checkbox As ICheckbox = Rx.CheckBox(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(Rx.CheckBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithDescription("The Description")))
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(Rx.CheckBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.
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
        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.
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
        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           WithLabel("Button", AddressOf GetLabel).
                           WithSuperTip("Super", AddressOf GetSuperTip).
                           WithDescription("Description", AddressOf GetDescription).
                           WithKeyTip("K2", AddressOf GetKeyTip).
                           OnCheck(False, AddressOf GetPressed, AddressOf OnButtonToggle,
                                   Sub(c) c.
                                       Do(Sub() Debug.WriteLine("On Toggle Fired")).
                                       ButFirst(Sub() Debug.WriteLine("Before Toggle Fired"))))


        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Visible, Not checkbox.Visible)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Enabled, Not checkbox.Enabled)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.IsChecked, Not checkbox.IsChecked)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Label, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.Description, String.Empty)
        Assert.That.PropertyMayBeModified(checkbox, Function(cb) cb.KeyTip, CType("A", KeyTip))

        Debug.WriteLine(checkbox)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.WithDescription("The Description").
                                                   WithLabel("The Label").
                                                   WithId("MyID"))

        Assert.That.SharedPropertiesAreEqual(control, Rx.Button(Sub(b) b.FromTemplate(control)))
    End Sub

    <TestMethod()>
    Public Sub EventCanFire()

        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.OnCheck(initialValue:=False, AddressOf GetPressed, AddressOf OnButtonToggle))

        Dim fired As Boolean = False

        AddHandler checkbox.Checked, Sub(sender, e) fired = True

        checkbox.IsChecked = True

        Assert.IsTrue(fired)

    End Sub

    <TestMethod()>
    Public Sub EventCanBeCancelled()
        Dim checkbox As ICheckbox = Rx.CheckBox(Sub(cbb) cbb.OnCheck(False, AddressOf GetPressed, AddressOf OnButtonToggle))

        AddHandler checkbox.Checking, Sub(sender, e) e.Cancel()

        checkbox.IsChecked = True

        Assert.IsFalse(checkbox.IsChecked)
    End Sub

    <TestMethod()>
    Public Sub HandlersCanBeAddedAndRemoved()

        Dim checkbox As ICheckbox = Rx.CheckBox()
        Dim beforeChange As EventHandler(Of CancelableEventArgs(Of Boolean)) = New EventHandler(Of CancelableEventArgs(Of Boolean))(Sub(sender, e) Return)
        Dim onChange As EventHandler(Of EventArgs(Of Boolean)) = New EventHandler(Of EventArgs(Of Boolean))(Sub(sender, e) Return)

        AddHandler checkbox.Checking, beforeChange
        AddHandler checkbox.Checked, onChange

        RemoveHandler checkbox.Checking, beforeChange
        RemoveHandler checkbox.Checked, onChange
    End Sub

End Class
﻿Imports RibbonX.Controls

<TestClass>
Public Class LabelControlTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As LabelControl = New LabelControl(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As LabelControl = New LabelControl(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyLabel())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New LabelControl())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As LabelControl = New LabelControl(config:=Sub(lcb) lcb.
                           Enabled().
                           Visible().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip"))

        Assert.IsTrue(control.Enabled)
        Assert.IsTrue(control.Visible)
        Assert.IsTrue(control.ShowLabel)
        Assert.AreEqual(control.Label, "The Label")
        Assert.AreEqual(control.ScreenTip, control.Label)
        Assert.AreEqual(control.SuperTip, "The Supertip")

        Debug.WriteLine(control)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyLabelII())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As LabelControl = BuildLabel()

        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(cb) cb.SuperTip, String.Empty)

        Debug.WriteLine(control)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As LabelControl = BuildReadonlyLabelII()

        Assert.That.SharedPropertiesAreEqual(control, New Button(template:=control))
    End Sub

    Public Shared Function BuildReadonlyLabel() As LabelControl
        Return New LabelControl(
            config:=Sub(lcb) lcb.
                        Enabled().
                        Visible().
                        ShowLabel().
                        WithLabel("Label").
                        WithScreenTip("Screentip").
                        WithSuperTip("Supertip"))
    End Function

    Public Shared Function BuildReadonlyLabelII() As LabelControl
        Return New LabelControl(
            config:=Sub(lcb) lcb.
                        Disabled().
                        Invisible().
                        WithLabel("Label").
                        HideLabel().
                        WithScreenTip("Screentip").
                        WithSuperTip("Supertip"))
    End Function

    Public Shared Function BuildLabel() As LabelControl
        Return New LabelControl(
            config:=Sub(lcb) lcb.
                        Enabled(AddressOf GetEnabledShared).
                        Visible(AddressOf GetVisibleShared).
                        WithLabel("Label", AddressOf GetLabelShared).
                        ShowLabel(AddressOf GetShowLabelShared).
                        WithScreenTip("Screentip", AddressOf GetScreenTipShared).
                        WithSuperTip("Supertip", AddressOf GetSuperTipShared))
    End Function

    Public Shared Function BuildLabelII() As LabelControl
        Return New LabelControl(
            config:=Sub(lcb) lcb.
                        Disabled(AddressOf GetEnabledShared).
                        Invisible(AddressOf GetVisibleShared).
                        WithLabel("Label", AddressOf GetLabelShared).
                        HideLabel(AddressOf GetShowLabelShared).
                        WithScreenTip("Screentip", AddressOf GetScreenTipShared).
                        WithSuperTip("Supertip", AddressOf GetSuperTipShared))
    End Function

End Class

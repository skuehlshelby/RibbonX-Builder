Imports System.Drawing
Imports RibbonX.Controls
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class ButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim button As Button = New Button(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim button As Button = New Button(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyButtonII())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Button())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As Button = New Button(Sub(bb) bb.
                           Visible().
                           Enabled().
                           ShowImage().
                           ShowLabel().
                           Large().
                           WithLabel("Button").
                           WithSuperTip("Super").
                           WithDescription("Description").
                           WithKeyTip("K2").
                           WithImage(Common.DollarSign))

        Assert.IsTrue(control.Enabled)
        Assert.IsTrue(control.Visible)
        Assert.IsTrue(control.ShowImage)
        Assert.IsTrue(control.ShowLabel)
        Assert.AreEqual(control.Size, ControlSize.Large)
        Assert.AreEqual(control.Label, "Button")
        Assert.AreEqual(control.ScreenTip, control.Label)
        Assert.AreEqual(control.SuperTip, "Super")
        Assert.AreEqual(control.Description, "Description")
        Assert.AreEqual(control.KeyTip, "K2")
        Assert.AreEqual(control.Image, Common.DollarSign)

        Debug.WriteLine(control)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyButton())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim button As Button = BuildButton()

        Assert.That.PropertyMayBeModified(button, Function(b) b.Visible, Not button.Visible)
        Assert.That.PropertyMayBeModified(button, Function(b) b.Enabled, Not button.Enabled)
        Assert.That.PropertyMayBeModified(button, Function(b) b.ShowLabel, Not button.ShowLabel)
        Assert.That.PropertyMayBeModified(button, Function(b) b.ShowImage, Not button.ShowImage)
        Assert.That.PropertyMayBeModified(button, Function(b) b.Label, String.Empty)
        Assert.That.PropertyMayBeModified(button, Function(b) b.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(button, Function(b) b.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(button, Function(b) b.Description, String.Empty)
        Assert.That.PropertyMayBeModified(button, Function(b) b.KeyTip, CType("P5", KeyTip))
        Assert.That.PropertyMayBeModified(button, Function(b) b.Size, ControlSize.Large)
        Assert.That.PropertyMayBeModified(button, Function(b) b.Image, RibbonImage.Create(Common.TraceError))

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As Button = BuildReadonlyButtonII()

        Assert.That.SharedPropertiesAreEqual(control, New EditBox(template:=control))
    End Sub

    <TestMethod()>
    Public Sub EventCanBeCancelled()
        Dim button As Button = New Button(Sub(bb) bb.RouteClickTo(AddressOf OnAction))

        AddHandler button.BeforeClick, Sub(sender, e) e.Cancel()

        Dim fired As Boolean = False

        AddHandler button.OnClick, Sub(sender, e) fired = True

        button.Click()

        Assert.IsFalse(fired)
    End Sub

    <TestMethod()>
    Public Sub EventCanFire()

        Dim button As Button = New Button(Sub(bb) bb.RouteClickTo(AddressOf OnAction))

        Dim fired As Boolean = False

        AddHandler button.OnClick, Sub(sender, e) fired = True

        button.Click()

        Assert.IsTrue(fired)

    End Sub

    <TestMethod()>
    Public Sub EventCanBeRemoved()
        Dim button As Button = New Button(Sub(bb) bb.RouteClickTo(AddressOf OnAction))
        Dim delegateWasInvoked As Boolean = False
        Dim callback As EventHandler = New EventHandler(Sub(o, e) delegateWasInvoked = True)
        Dim callback2 As EventHandler(Of CancelableEventArgs) = New EventHandler(Of CancelableEventArgs)(Sub(o, e) delegateWasInvoked = True)

        AddHandler button.OnClick, callback
        AddHandler button.BeforeClick, callback2

        button.Click()

        Assert.IsTrue(delegateWasInvoked)

        delegateWasInvoked = False

        RemoveHandler button.OnClick, callback
        RemoveHandler button.BeforeClick, callback2

        button.Click()

        Assert.IsFalse(delegateWasInvoked)
    End Sub

    <TestMethod()>
    Public Sub ClonedControl_Click_NoThrow() 'This was actually a problem at one point
        Dim toggle As ToggleButton = New ToggleButton()
        Dim button As Button = New Button(template:=toggle)

        button.Click()
    End Sub

    Public Shared Function BuildReadonlyButton() As Button
        Return New Button(Sub(bb) bb.
                   Invisible().
                   Disabled().
                   HideImage().
                   HideLabel().
                   Normal().
                   WithLabel("Button").
                   WithSuperTip("Super").
                   WithDescription("Description").
                   WithKeyTip("K2").
                   WithImage("TheNameOfMyCachedImage"))
    End Function

    Public Shared Function BuildReadonlyButtonII() As Button
        Return New Button(Sub(bb) bb.
                   Visible().
                   Enabled().
                   ShowImage().
                   ShowLabel().
                   Large().
                   WithLabel("Button").
                   WithScreenTip("Button").
                   WithSuperTip("Super").
                   WithDescription("Description").
                   WithKeyTip("K2").
                   WithImage("TheNameOfMyCachedImage"))
    End Function

    Public Shared Function BuildButton() As Button
        Return New Button(Sub(bb) bb.
                           Visible(AddressOf GetVisibleShared).
                           Enabled(AddressOf GetEnabledShared).
                           Normal(AddressOf GetSizeShared).
                           WithLabel("Button", AddressOf GetLabelShared).
                           ShowLabel(AddressOf GetShowLabelShared).
                           WithScreenTip("Button", AddressOf GetScreenTipShared).
                           WithSuperTip("Super", AddressOf GetSuperTipShared).
                           WithDescription("Description", AddressOf GetDescriptionShared).
                           WithKeyTip("K2", AddressOf GetKeyTipShared).
                           WithImage(New Bitmap(48, 48), AddressOf GetImageShared).
                           ShowImage(AddressOf GetShowImageShared).
                           RouteClickTo(AddressOf OnActionShared).
                           BeforeClick(Sub() Debug.WriteLine("Before Click Fired")).
                           OnClick(Sub() Debug.WriteLine("On Click Fired")))
    End Function

    Public Shared Function BuildButtonII() As Button
        Return New Button(Sub(bb) bb.
                           Invisible(AddressOf GetVisibleShared).
                           Disabled(AddressOf GetEnabledShared).
                           Large(AddressOf GetSizeShared).
                           WithLabel("Button", AddressOf GetLabelShared, False).
                           WithScreenTip("Button", AddressOf GetScreenTipShared).
                           HideLabel(AddressOf GetShowLabelShared).
                           WithSuperTip("Super", AddressOf GetSuperTipShared).
                           WithDescription("Description", AddressOf GetDescriptionShared).
                           WithKeyTip("K2", AddressOf GetKeyTipShared).
                           WithImage(New Bitmap(48, 48), AddressOf GetImageShared).
                           HideImage(AddressOf GetShowImageShared).
                           RouteClickTo(AddressOf OnActionShared).
                           BeforeClick(Sub() Debug.WriteLine("Before Click Fired")).
                           OnClick(Sub() Debug.WriteLine("On Click Fired")))
    End Function

End Class
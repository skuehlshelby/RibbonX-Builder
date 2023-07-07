Imports System.Drawing
Imports RibbonX
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class ButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim button As IButton = RxApi.Button(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim button As IButton = RxApi.Button()
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyButtonII())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.Button())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As IButton = RxApi.Button(Sub(bb) bb.
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
        Dim button As IButton = BuildButton()

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
        Dim control As IButton = BuildReadonlyButtonII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.EditBox(Sub(b) b.FromTemplate(control)))
    End Sub

    <TestMethod()>
    Public Sub EventCanBeCancelled()
        Dim button As IButton = RxApi.Button(Sub(bb) bb.OnClick(AddressOf OnAction))

        AddHandler button.Clicking, Sub(sender, e) e.Cancel()

        Dim fired As Boolean = False

        AddHandler button.Clicked, Sub(sender, e) fired = True

        button.Click()

        Assert.IsFalse(fired)
    End Sub

    <TestMethod()>
    Public Sub EventCanFire()

        Dim button As IButton = RxApi.Button(Sub(bb) bb.OnClick(AddressOf OnAction))

        Dim fired As Boolean = False

        AddHandler button.Clicked, Sub(sender, e) fired = True

        button.Click()

        Assert.IsTrue(fired)

    End Sub

    <TestMethod()>
    Public Sub EventCanBeRemoved()
        Dim button As IButton = RxApi.Button()
        Dim delegateWasInvoked As Boolean = False
        Dim callback As EventHandler = New EventHandler(Sub(o, e) delegateWasInvoked = True)
        Dim callback2 As EventHandler(Of CancelableEventArgs) = New EventHandler(Of CancelableEventArgs)(Sub(o, e) delegateWasInvoked = True)

        AddHandler button.Clicked, callback
        AddHandler button.Clicking, callback2

        button.Click()

        Assert.IsTrue(delegateWasInvoked)

        delegateWasInvoked = False

        RemoveHandler button.Clicked, callback
        RemoveHandler button.Clicking, callback2

        button.Click()

        Assert.IsFalse(delegateWasInvoked)
    End Sub

    <TestMethod()>
    Public Sub ClonedControl_Click_NoThrow() 'This was actually a problem at one point
        Dim toggle As IToggleButton = RxApi.ToggleButton()
        Dim button As IButton = RxApi.Button(Sub(b) b.FromTemplate(toggle))

        button.Click()
    End Sub

    Friend Shared Function BuildReadonlyButton() As IButton
        Return RxApi.Button(Sub(bb) bb.
                   Invisible().
                   Disabled().
                   HideImage().
                   HideLabel().
                   Normal().
                   WithLabel("Button").
                   WithSuperTip("Super").
                   WithDescription("Description").
                   WithKeyTip("K2").
                   WithImage(Common.DollarSign))
    End Function

    Friend Shared Function BuildReadonlyButtonII() As IButton
        Return RxApi.Button(Sub(bb) bb.
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
                               WithImage(Common.DollarSign))
    End Function

    Friend Shared Function BuildButton() As IButton
        Return RxApi.Button(Sub(bb) bb.
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
                           OnClick(AddressOf OnActionShared,
            Sub(b) b.
                                                   Do(Sub() Debug.WriteLine("On Click Fired")).
                                                   ButFirst(Sub() Debug.WriteLine("Before Click Fired"))))
    End Function

    Friend Shared Function BuildButtonII() As IButton
        Return RxApi.Button(Sub(bb) bb.
                           Invisible(AddressOf GetVisibleShared).
                           Disabled(AddressOf GetEnabledShared).
                           Large(AddressOf GetSizeShared).
                           WithLabel("Button", AddressOf GetLabelShared, False).
                           WithScreenTip("Button", AddressOf GetScreenTipShared).
                           HideLabel(AddressOf GetShowLabelShared).
                           WithSuperTip("Super", AddressOf GetSuperTipShared).
                           WithDescription("Description", AddressOf GetDescriptionShared).
                           WithKeyTip("K2", AddressOf GetKeyTipShared).
                           WithImage(Common.DollarSign, AddressOf GetBuiltInImageShared).
                           HideImage(AddressOf GetShowImageShared).
                           OnClick(AddressOf OnActionShared,
            Sub(b) b.
                                                   Do(Sub() Debug.WriteLine("On Click Fired")).
                                                   ButFirst(Sub() Debug.WriteLine("Before Click Fired"))))
    End Function

End Class
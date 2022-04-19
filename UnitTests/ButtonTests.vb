Imports System.Drawing
Imports RibbonFactory
Imports RibbonFactory.Controls
Imports RibbonFactory.Controls
Imports RibbonFactory.Controls.Events
Imports RibbonFactory.Enums
Imports RibbonFactory.Enums.ImageMSO

<TestClass()>
Public Class ButtonTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim button As Button = New Button(Sub(bb) bb.Visible(), Nothing)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim button As Button = New Button(Nothing, Nothing)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Dim ribbon As Ribbon = CreateSimpleRibbon(
            New Button(Sub(bb) bb.
                           Visible().
                           Enabled().
                           ShowImage().
                           ShowLabel().
                           WithLabel("Button").
                           WithSuperTip("Super").
                           WithImage(Common.DollarSign)))

        Assert.IsTrue(ribbon.GetErrors().None)

        Debug.WriteLine(ribbon.RibbonX)
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Dim button As Button = New Button()

        ContainsNoNullValuesByDefaultHelper(button)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim button As Button = New Button(Sub(bb) bb.
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

        Assert.IsTrue(button.Enabled)
        Assert.IsTrue(button.Visible)
        Assert.IsTrue(button.ShowImage)
        Assert.IsTrue(button.ShowLabel)
        Assert.AreEqual(button.Size, ControlSize.large)
        Assert.AreEqual(button.Label, "Button")
        Assert.AreEqual(button.ScreenTip, button.Label)
        Assert.AreEqual(button.SuperTip, "Super")
        Assert.AreEqual(button.Description, "Description")
        Assert.AreEqual(button.KeyTip, "K2")
        Assert.AreEqual(button.Image, Common.DollarSign)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim button As Button = New Button(Sub(bb) bb.
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

        Assert.ThrowsException(Of Exception)(Sub() button.Visible = True)
        Assert.ThrowsException(Of Exception)(Sub() button.Enabled = True)
        Assert.ThrowsException(Of Exception)(Sub() button.ShowLabel = True)
        Assert.ThrowsException(Of Exception)(Sub() button.ShowImage = True)
        Assert.ThrowsException(Of Exception)(Sub() button.Label = "New Label")
        Assert.ThrowsException(Of Exception)(Sub() button.ScreenTip = "New Screentip")
        Assert.ThrowsException(Of Exception)(Sub() button.SuperTip = "New Supertip")
        Assert.ThrowsException(Of Exception)(Sub() button.KeyTip = "K5")
        Assert.ThrowsException(Of Exception)(Sub() button.Description = "New Description")
        Assert.ThrowsException(Of Exception)(Sub() button.Image = RibbonImage.Create(Common.SadFace))
        Assert.ThrowsException(Of Exception)(Sub() button.Size = ControlSize.large)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim button As Button = New Button(Sub(bb) bb.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           Normal(AddressOf GetSize).
                           WithLabel("Button", AddressOf GetLabel).
                           HideLabel(AddressOf GetShowLabel).
                           WithSuperTip("Super", AddressOf GetSuperTip).
                           WithDescription("Description", AddressOf GetDescription).
                           WithKeyTip("K2", AddressOf GetKeyTip).
                           WithImage(New Bitmap(48, 48), AddressOf GetImage).
                           ShowImage(AddressOf GetShowImage).
                           RouteClickTo(AddressOf OnAction).
                           BeforeClick(Sub() Debug.WriteLine("Before Click Fired")).
                           OnClick(Sub() Debug.WriteLine("On Click Fired")))

        Assert.IsTrue(button.Visible)
        button.Visible = False
        Assert.IsFalse(button.Visible)

        Assert.IsTrue(button.Enabled)
        button.Enabled = False
        Assert.IsFalse(button.Enabled)

        Assert.IsFalse(button.ShowLabel)
        button.ShowLabel = True
        Assert.IsTrue(button.ShowLabel)

        Assert.IsTrue(button.ShowImage)
        button.ShowImage = False
        Assert.IsFalse(button.ShowImage)

        Assert.AreEqual(button.Label, "Button")
        button.Label = "New Label"
        Assert.AreEqual(button.Label, "New Label")

        Assert.AreEqual(button.ScreenTip, "Button")
        button.ScreenTip = "New ScreenTip"
        Assert.AreEqual(button.ScreenTip, "New ScreenTip")

        Assert.AreEqual(button.SuperTip, "Super")
        button.SuperTip = "New SuperTip"
        Assert.AreEqual(button.SuperTip, "New SuperTip")

        Assert.AreEqual(button.Description, "Description")
        button.Description = "New Description"
        Assert.AreEqual(button.Description, "New Description")

        Assert.AreEqual(button.KeyTip, "K2")
        button.KeyTip = "A"
        Assert.AreEqual(button.KeyTip, "A")

        Assert.AreEqual(button.Size, ControlSize.normal)
        button.Size = ControlSize.large
        Assert.AreEqual(button.Size, ControlSize.large)

        button.Image = RibbonImage.Create(Common.TraceError)
        Assert.AreEqual(button.Image, Common.TraceError)

        Debug.WriteLine(button)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim firstButton As Button = New Button(Sub(bb) bb.WithDescription("The Description").WithLabel("The Label").WithId("MyID"))
        Dim secondButton As Button = New Button(firstButton)

        Assert.AreEqual(secondButton.Description, "The Description")
        Assert.AreEqual(secondButton.Label, "The Label")
        Assert.AreNotEqual(secondButton.ID, firstButton.ID)
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
        Dim button As Button = New Button(toggle)

        button.Click()
    End Sub

End Class
Imports RibbonX.Builders
Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class DropDownTests
    Inherits TestBase

    Private Const LABEL As String = "My Dropdown"
    Private Const SCREENTIP As String = "My Dropdown Screentip"
    Private Const SUPERTIP As String = "More Info About My Dropdown"
    Private Const KEYTIP As String = "J"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim dropdown As DropDown = New DropDown(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim dropdown As DropDown = New DropDown(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyDropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New DropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim dropDown As DropDown = BuildReadonlyDropDown()

        Assert.IsTrue(dropDown.Enabled)
        Assert.IsTrue(dropDown.Visible)
        Assert.IsTrue(dropDown.ShowImage)
        Assert.IsTrue(dropDown.ShowLabel)
        Assert.AreEqual(dropDown.Label, LABEL)
        Assert.AreEqual(dropDown.ScreenTip, SCREENTIP)
        Assert.AreEqual(dropDown.SuperTip, SUPERTIP)
        Assert.AreEqual(dropDown.KeyTip.ToString(), KEYTIP)
        Assert.AreEqual(dropDown.Image, Common.FileFind)

        Debug.WriteLine(dropDown)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyDropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim dropdown As DropDown = BuildDropDownII()

        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.Visible, Not dropdown.Visible)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.Enabled, Not dropdown.Enabled)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.ShowLabel, Not dropdown.ShowLabel)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.ShowImage, Not dropdown.ShowImage)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.Label, String.Empty)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.KeyTip, CType("P5", KeyTip))
        Assert.That.PropertyMayBeModified(dropdown, Function(b) b.Image, RibbonImage.Create(Common.TraceError))

        Debug.WriteLine(dropdown)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim dropDown As DropDown = BuildDropDown()

        Assert.That.SharedPropertiesAreEqual(dropDown, New CheckBox(template:=dropDown))
    End Sub

    <TestMethod>
    Public Sub CanAddItems()
        Dim dropDown As DropDown = New DropDown() From {
            Item.Blank,
            Item.Blank,
            Item.Blank,
            Item.Blank
        }

        Assert.AreEqual(dropDown.Count, 4)
    End Sub

    <TestMethod>
    Public Sub CanRemoveItems()
        Dim items As Item() = {
            New Item(Sub(ib) ib.WithLabel("Item1")),
            New Item(Sub(ib) ib.WithLabel("Item2")),
            New Item(Sub(ib) ib.WithLabel("Item3")),
            New Item(Sub(ib) ib.WithLabel("Item4"))
        }

        Dim dropDown As DropDown = BuildReadonlyDropDownII()

        Array.ForEach(items, Sub(item) dropDown.Add(item))

        Array.ForEach(items, Sub(item) dropDown.Remove(item))
    End Sub

    <TestMethod>
    Public Sub CanClearItems()
        Dim items As Item() = {
            New Item(Sub(ib) ib.WithLabel("Item1")),
            New Item(Sub(ib) ib.WithLabel("Item2")),
            New Item(Sub(ib) ib.WithLabel("Item3")),
            New Item(Sub(ib) ib.WithLabel("Item4"))
        }

        Dim dropDown As DropDown = BuildReadonlyDropDownII()

        Array.ForEach(items, Sub(item) dropDown.Add(item))

        Assert.AreEqual(dropDown.Count, 4)

        dropDown.Clear()

        Assert.AreEqual(dropDown.Count, 0)
    End Sub

    <TestMethod>
    Public Sub SelectedItemIsBlankWhenDropdownIsEmpty()
        Assert.AreEqual(BuildDropDownII().Selected, Item.Blank)
    End Sub

    <TestMethod>
    Public Sub PropertyChangeTriggersRefresh()
        Dim control As DropDown = New DropDown(Sub(cbb) cbb.WithLabel("The Label", AddressOf GetLabel))

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(cb) cb.Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub ItemAddTriggersRefresh()
        Dim control As DropDown = New DropDown()

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(cb) cb.Add(Item.Blank()))
    End Sub

    <TestMethod>
    Public Sub ChangingChildItemTriggersRefresh()
        Dim control As DropDown = New DropDown() From {
            New Item(Sub(ib) ib.WithLabel("Item1"))
        }

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(c) c.First().Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub SelectionCanBeChanged()
        Dim items As Item() = {
            New Item(Sub(ib) ib.WithLabel("Item1")),
            New Item(Sub(ib) ib.WithLabel("Item2")),
            New Item(Sub(ib) ib.WithLabel("Item3")),
            New Item(Sub(ib) ib.WithLabel("Item4"))
        }

        Dim control As DropDown = BuildDropDown()

        Array.ForEach(items, Sub(item) control.Add(item))

        For i As Integer = 0 To control.Count - 1
            control.Selected = items(i)
            Assert.AreEqual(items(i), control.Selected)
            Assert.AreEqual(i, control.SelectedItemIndex)
        Next

        For i As Integer = 0 To control.Count - 1
            control.SelectedItemIndex = i
            Assert.AreEqual(items(i), control.Selected)
            Assert.AreEqual(i, control.SelectedItemIndex)
        Next
    End Sub

    Public Shared Function BuildReadonlyDropDown() As DropDown
        Return New DropDown(config:=Sub(bb As IDropdownBuilder) bb.
                           InsertAfter(Excel.AccessibilityChecker).
                           Visible().
                           Enabled().
                           ShowImage().
                           WithImage(Common.FileFind).
                           WithLabel(LABEL).
                           ShowLabel().
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           WithKeyTip(KEYTIP),
                            buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()})
    End Function

    Public Shared Function BuildReadonlyDropDownII() As DropDown
        Return New DropDown(config:=Sub(bb As IDropdownBuilder) bb.
                           InsertBefore(Excel.AccessibilityChecker).
                           Invisible().
                           Disabled().
                           WithImage(Common.FileFind).
                           HideImage().
                           WithLabel(LABEL).
                           HideLabel().
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           WithKeyTip(KEYTIP),
                            buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()})
    End Function

    Public Shared Function BuildDropDown() As DropDown
        Return New DropDown(config:=Sub(bb As IDropdownBuilder) bb.
                           Visible(AddressOf GetVisibleShared).
                           Enabled(AddressOf GetEnabledShared).
                           ShowImage(AddressOf GetShowImageShared).
                           WithImage(BlankBitmap(), AddressOf GetImageShared).
                           WithLabel(LABEL, AddressOf GetLabelShared).
                           ShowLabel(AddressOf GetShowLabelShared).
                           WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                           WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                           WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                           GetItemCountFrom(AddressOf GetItemCountShared).
                           GetSelectedItemIndexFrom(AddressOf GetSelectedItemIndexShared, AddressOf OnSelectionChangeShared).
                           GetItemIdFrom(AddressOf GetItemIDShared).
                           GetItemImageFrom(AddressOf GetItemImageShared).
                           GetItemLabelFrom(AddressOf GetItemLabelShared).
                           GetItemScreenTipFrom(AddressOf GetItemScreenTipShared).
                           GetItemSuperTipFrom(AddressOf GetItemSuperTipShared),
                            buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()})
    End Function

    Public Shared Function BuildDropDownII() As DropDown
        Return New DropDown(config:=Sub(bb As IDropdownBuilder) bb.
                           Visible(AddressOf GetVisibleShared).
                           Enabled(AddressOf GetEnabledShared).
                           ShowImage(AddressOf GetShowImageShared).
                           WithImage(BlankBitmap(), AddressOf GetImageShared).
                           WithLabel(LABEL, AddressOf GetLabelShared).
                           ShowLabel(AddressOf GetShowLabelShared).
                           WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                           WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                           WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                           GetItemCountFrom(AddressOf GetItemCountShared).
                           GetSelectedItemIndexFrom(AddressOf GetSelectedItemIndexShared, AddressOf OnSelectionChangeShared).
                           GetItemIdFrom(AddressOf GetItemIDShared).
                           GetItemImageFrom(AddressOf GetItemImageShared).
                           GetItemLabelFrom(AddressOf GetItemLabelShared).
                           GetItemScreenTipFrom(AddressOf GetItemScreenTipShared).
                           GetItemSuperTipFrom(AddressOf GetItemSuperTipShared),
                            buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()})
    End Function

End Class
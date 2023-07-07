Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes
Imports RibbonX
Imports Rx = RibbonX.RxApi

<TestClass()>
Public Class DropDownTests
    Inherits TestBase

    Private Const LABEL As String = "My Dropdown"
    Private Const SCREENTIP As String = "My Dropdown Screentip"
    Private Const SUPERTIP As String = "More Info About My Dropdown"
    Private Const KEYTIP As String = "J"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim dropdown As IDropDown = Rx.DropDown(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim dropdown As IDropDown = Rx.DropDown(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildReadonlyDropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(Rx.DropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim dropDown As IDropDown = BuildReadonlyDropDown()

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
        Dim dropdown As IDropDown = BuildDropDownII()

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
        Dim dropDown As IDropDown = BuildDropDown()

        Assert.That.SharedPropertiesAreEqual(dropDown, Rx.CheckBox(Sub(b) b.FromTemplate(dropDown)))
    End Sub

    <TestMethod>
    Public Sub CanAddItems()
        Dim dropDown As IDropDown = Rx.DropDown()
        dropDown.Add(Rx.Item(), Rx.Item(), Rx.Item(), Rx.Item())

        Assert.AreEqual(dropDown.Count, 4)
    End Sub

    <TestMethod>
    Public Sub CanRemoveItems()
        Dim items As IItem() = {
            Rx.Item(Sub(b) b.WithLabel("Item1")),
            Rx.Item(Sub(b) b.WithLabel("Item2")),
            Rx.Item(Sub(b) b.WithLabel("Item3")),
            Rx.Item(Sub(b) b.WithLabel("Item4"))
        }

        Dim dropDown As IDropDown = BuildReadonlyDropDownII()

        Array.ForEach(items, Sub(item) dropDown.Add(item))

        Array.ForEach(items, Sub(item) dropDown.Remove(item))
    End Sub

    <TestMethod>
    Public Sub CanClearItems()
        Dim items As IItem() = {
            Rx.Item(Sub(b) b.WithLabel("Item1")),
            Rx.Item(Sub(b) b.WithLabel("Item2")),
            Rx.Item(Sub(b) b.WithLabel("Item3")),
            Rx.Item(Sub(b) b.WithLabel("Item4"))
        }

        Dim dropDown As IDropDown = BuildReadonlyDropDownII()

        Array.ForEach(items, Sub(item) dropDown.Add(item))

        Assert.AreEqual(dropDown.Count, 4)

        dropDown.Clear()

        Assert.AreEqual(dropDown.Count, 0)
    End Sub

    <TestMethod>
    Public Sub SelectedItemIsBlankWhenDropdownIsEmpty()
        Dim blank As IItem = Rx.Item()
        Dim dropdown As IDropDown = Rx.DropDown(Sub(b) b.WithBlank(blank))
        Assert.AreEqual(dropdown.Selected, blank)
    End Sub

    <TestMethod>
    Public Sub PropertyChangeTriggersRefresh()
        Dim control As IDropDown = Rx.DropDown(Sub(cbb) cbb.WithLabel("The Label", AddressOf GetLabel))

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(cb) cb.Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub ItemAddTriggersRefresh()
        Dim control As IDropDown = Rx.DropDown()

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(cb) cb.Add(Rx.Item()))
    End Sub

    <TestMethod>
    Public Sub ChangingChildItemTriggersRefresh()
        Dim control As IDropDown = Rx.DropDown()
        control.Add(Rx.Item(Sub(b) b.WithLabel("Item1")))

        Assert.That.ValueChangedIsRaisedOnce(control, Sub(c) c.First().Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub SelectionCanBeChanged()
        Dim items As IItem() = {
            Rx.Item(Sub(b) b.WithLabel("Item1")),
            Rx.Item(Sub(b) b.WithLabel("Item2")),
            Rx.Item(Sub(b) b.WithLabel("Item3")),
            Rx.Item(Sub(b) b.WithLabel("Item4"))
        }

        Dim control As IDropDown = BuildDropDown()

        control.Add(items)

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

    Public Shared Function BuildReadonlyDropDown() As IDropDown
        Return Rx.DropDown(Sub(bb) bb.
                           InsertAfter(Excel.AccessibilityChecker).
                           Visible().
                           Enabled().
                           ShowImage().
                           WithImage(Common.FileFind).
                           WithLabel(LABEL).
                           ShowLabel().
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           WithKeyTip(KEYTIP).
                           WithButtons(ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()))
    End Function

    Public Shared Function BuildReadonlyDropDownII() As IDropDown
        Return Rx.DropDown(Sub(bb) bb.
                           InsertBefore(Excel.AccessibilityChecker).
                           Invisible().
                           Disabled().
                           WithImage(Common.FileFind).
                           HideImage().
                           WithLabel(LABEL).
                           HideLabel().
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           WithKeyTip(KEYTIP).
                           WithButtons(ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()))
    End Function

    Public Shared Function BuildDropDown() As IDropDown
        Return Rx.DropDown(Sub(bb) bb.
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
                           OnSelectionChange(AddressOf GetSelectedItemIndexShared, AddressOf OnSelectionChangeShared).
                           GetItemIdFrom(AddressOf GetItemIDShared).
                           GetItemImageFrom(AddressOf GetItemImageShared).
                           GetItemLabelFrom(AddressOf GetItemLabelShared).
                           GetItemScreenTipFrom(AddressOf GetItemScreenTipShared).
                           GetItemSuperTipFrom(AddressOf GetItemSuperTipShared).
                           WithButtons(ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()))
    End Function

    Public Shared Function BuildDropDownII() As IDropDown
        Return Rx.DropDown(Sub(bb) bb.
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
                           OnSelectionChange(AddressOf GetSelectedItemIndexShared, AddressOf OnSelectionChangeShared).
                           GetItemIdFrom(AddressOf GetItemIDShared).
                           GetItemImageFrom(AddressOf GetItemImageShared).
                           GetItemLabelFrom(AddressOf GetItemLabelShared).
                           GetItemScreenTipFrom(AddressOf GetItemScreenTipShared).
                           GetItemSuperTipFrom(AddressOf GetItemSuperTipShared).
                           WithButtons(ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()))
    End Function

End Class
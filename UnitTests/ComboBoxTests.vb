
Imports RibbonX
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class IComboBoxTests
    Inherits TestBase

    Private Const LABEL As String = "My IComboBox"
    Private Const SCREENTIP As String = "My IComboBox Screentip"
    Private Const SUPERTIP As String = "More Info About My IComboBox"
    Private Const TEXT As String = "The Text"
    Private Const KEYTIP As String = "E"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim combobox As IComboBox = RxApi.ComboBox(Sub(b) b.FromTemplate(Nothing))
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim combobox As IComboBox = RxApi.ComboBox(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildComboBox())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(RxApi.ComboBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim combobox As IComboBox = BuildReadonlyComboBox()

        Assert.IsTrue(combobox.Enabled)
        Assert.IsTrue(combobox.Visible)
        Assert.IsTrue(combobox.ShowLabel)
        Assert.IsTrue(combobox.ShowImage)
        Assert.IsFalse(combobox.IsReadOnly)
        Assert.AreEqual(combobox.Label, LABEL)
        Assert.AreEqual(combobox.ScreenTip, SCREENTIP)
        Assert.AreEqual(combobox.SuperTip, SUPERTIP)
        Assert.AreEqual(combobox.MaxLength, CByte(40))
        Assert.AreEqual(combobox.Image, Common.SadFace)
        Assert.AreEqual(combobox.KeyTip.ToString(), KEYTIP)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyComboBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As IComboBox = BuildComboBoxII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.KeyTip, CType("R", KeyTip))
        Assert.That.PropertyMayBeModified(control, Function(c) c.Text, String.Empty)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As IComboBox = BuildReadonlyComboBoxII()

        Assert.That.SharedPropertiesAreEqual(control, RxApi.DropDown(Sub(b) b.FromTemplate(control)))
    End Sub

    <TestMethod>
    Public Sub CanAddItemsAfterIComboBoxCreation()
        Dim combobox As IComboBox = RxApi.ComboBox()

        For Index As Integer = 1 To 4
            combobox.Add(RxApi.Item())
        Next

        Assert.AreEqual(combobox.Count, 4)
    End Sub

    <TestMethod>
    Public Sub CanClearItems()
        Dim combobox As IComboBox = RxApi.ComboBox()

        For Index As Integer = 1 To 4
            combobox.Add(RxApi.Item())
        Next

        Assert.AreEqual(combobox.Count, 4)

        combobox.Clear()

        Assert.AreEqual(combobox.Count, 0)
    End Sub

    <TestMethod>
    Public Sub CanRemoveItems()
        Dim combobox As IComboBox = RxApi.ComboBox()

        For Index As Integer = 1 To 4
            combobox.Add(RxApi.Item())
        Next

        Assert.AreEqual(combobox.Count, 4)

        While combobox.Any()
            Assert.IsTrue(combobox.Remove(combobox.First()))
        End While

        Assert.IsFalse(combobox.Remove(RxApi.Item()))

        Assert.AreEqual(combobox.Count, 0)
    End Sub

    <TestMethod>
    Public Sub PropertyChangeTriggersRefresh()
        Dim combobox As IComboBox = RxApi.ComboBox(Sub(cbb) cbb.WithLabel("The Label", AddressOf GetLabel))

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub ItemAddTriggersRefresh()
        Dim combobox As IComboBox = RxApi.ComboBox()

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.Add(RxApi.Item()))
    End Sub

    <TestMethod>
    Public Sub ChangingChildItemTriggersRefresh()
        Dim combobox As IComboBox = RxApi.ComboBox()

        combobox.Add(RxApi.Item())

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.First().Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub EventCanBeSubscribedTo()
        Dim combobox As IComboBox = RxApi.ComboBox()

        combobox.Add(RxApi.Item())

        Dim before As EventHandler(Of CancelableEventArgs(Of String)) = Sub(s, e) Return
        Dim onChange As EventHandler(Of EventArgs(Of String)) = Sub(s, e) Return

        AddHandler combobox.Changing, before
        AddHandler combobox.Changed, onChange

        RemoveHandler combobox.Changing, before
        RemoveHandler combobox.Changed, onChange
    End Sub

    Private NotInheritable Class TestItemTemplate
        Implements IItemTemplate

        Public Function Match(obj As Object) As Boolean Implements IItemTemplate.Match
            Return True
        End Function

        Public Function Apply(obj As Object) As IItem Implements IItemTemplate.Apply
            Dim type As Type = obj.GetType()

            Return RxApi.Item(Sub(b) b.WithId(type.Name).WithLabel(type.Name).WithSuperTip(type.FullName).WithTag(type))
        End Function
    End Class

    <TestMethod>
    Public Sub TemplateCanBeAdded()
        Dim combobox As IComboBox = RxApi.ComboBox()

        combobox.AddTemplate(New TestItemTemplate())
    End Sub

    <TestMethod>
    Public Sub TemplateCanBeUsedToAddItems()
        Dim combobox As IComboBox = RxApi.ComboBox()

        combobox.AddTemplate(New TestItemTemplate())

        combobox.AddTemplatedItem(New Exception)
        combobox.AddTemplatedItem(New Boolean())

        Assert.AreEqual(2, combobox.Count)
    End Sub

    <TestMethod>
    Public Sub TemplateCanBeUsedToRemoveItems()
        Dim combobox As IComboBox = RxApi.ComboBox()

        combobox.AddTemplate(New TestItemTemplate())

        combobox.AddTemplatedItem(New Exception)
        combobox.AddTemplatedItem(New Boolean())

        Assert.AreEqual(2, combobox.Count)

        combobox.RemoveTemplatedItem(New Exception)
        combobox.RemoveTemplatedItem(New Boolean())

        Assert.AreEqual(0, combobox.Count)
    End Sub

    Public Shared Function BuildReadonlyComboBox() As IComboBox
        Return RxApi.ComboBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel(LABEL).
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           ShowLabel().
                           ShowImage().
                           ShowItemImages().
                           WithKeyTip(KEYTIP).
                           WithMaximumInputLength(40).
                           WithImage(Common.SadFace))
    End Function

    Public Shared Function BuildReadonlyComboBoxII() As IComboBox
        Return RxApi.ComboBox(Sub(cbb) cbb.
                           Invisible().
                           Disabled().
                           WithLabel(LABEL).
                           WithScreenTip(SCREENTIP).
                           WithSuperTip(SUPERTIP).
                           HideLabel().
                           HideImage().
                           HideItemImages().
                           WithKeyTip(KEYTIP).
                           WithMaximumInputLength(40).
                           WithImage(Common.SadFace))
    End Function

    Public Shared Function BuildComboBox() As IComboBox
        Return RxApi.ComboBox(Sub(cbb) cbb.
                           Visible(AddressOf GetVisibleShared).
                           Enabled(AddressOf GetEnabledShared).
                           WithLabel(LABEL, AddressOf GetLabelShared).
                           WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                           WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                           ShowLabel(AddressOf GetShowLabelShared).
                           ShowImage(AddressOf GetShowImageShared).
                           ShowItemImages().
                           WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                           WithMaximumInputLength(40).
                           WithText(TEXT, AddressOf GetTextShared, AddressOf OnChangeShared).
                           WithImage(BlankBitmap(), AddressOf GetImageShared))
    End Function

    Public Shared Function BuildComboBoxII() As IComboBox
        Return RxApi.ComboBox(Sub(cbb) cbb.
                           Invisible(AddressOf GetVisibleShared).
                           Disabled(AddressOf GetEnabledShared).
                           WithLabel(LABEL, AddressOf GetLabelShared).
                           WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                           WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                           HideLabel(AddressOf GetShowLabelShared).
                           HideImage(AddressOf GetShowImageShared).
                           HideItemImages().
                           WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                           WithMaximumInputLength(40).
                           WithText(TEXT, AddressOf GetTextShared, AddressOf OnChangeShared).
                           WithImage(BlankBitmap(), AddressOf GetImageShared))
    End Function

End Class
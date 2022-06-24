Imports RibbonX.Controls
Imports RibbonX.Controls.EventArgs
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass()>
Public Class ComboBoxTests
    Inherits TestBase

    Private Const LABEL As String = "My ComboBox"
    Private Const SCREENTIP As String = "My ComboBox Screentip"
    Private Const SUPERTIP As String = "More Info About My ComboBox"
    Private Const TEXT As String = "The Text"
    Private Const KEYTIP As String = "E"

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim combobox As ComboBox = New ComboBox(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim combobox As ComboBox = New ComboBox(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildComboBox())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New ComboBox())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim combobox As ComboBox = BuildReadonlyComboBox()

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
        Dim control As ComboBox = BuildComboBoxII()

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
        Dim control As ComboBox = BuildReadonlyComboBoxII()

        Assert.That.SharedPropertiesAreEqual(control, New DropDown(template:=control))
    End Sub

    <TestMethod>
    Public Sub CanAddItemsAfterComboBoxCreation()
        Dim combobox As ComboBox = New ComboBox() From {
            Item.Blank(),
            Item.Blank(),
            Item.Blank(),
            Item.Blank()
        }

        Assert.AreEqual(combobox.Count, 4)
    End Sub

    <TestMethod>
    Public Sub CanClearItems()
        Dim combobox As ComboBox = New ComboBox() From {
            Item.Blank(),
            Item.Blank(),
            Item.Blank(),
            Item.Blank()
        }

        Assert.AreEqual(combobox.Count, 4)

        combobox.Clear()

        Assert.AreEqual(combobox.Count, 0)
    End Sub

    <TestMethod>
    Public Sub CanRemoveItems()
        Dim combobox As ComboBox = New ComboBox() From {
            Item.Blank(), Item.Blank(), Item.Blank()
        }

        Assert.AreEqual(combobox.Count, 3)

        While combobox.Any()
            Assert.IsTrue(combobox.Remove(combobox.First()))
        End While

        Assert.IsFalse(combobox.Remove(Item.Blank()))

        Assert.AreEqual(combobox.Count, 0)
    End Sub

    <TestMethod>
    Public Sub PropertyChangeTriggersRefresh()
        Dim combobox As ComboBox = New ComboBox(Sub(cbb) cbb.WithLabel("The Label", AddressOf GetLabel))

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub ItemAddTriggersRefresh()
        Dim combobox As ComboBox = New ComboBox()

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.Add(Item.Blank()))
    End Sub

    <TestMethod>
    Public Sub ChangingChildItemTriggersRefresh()
        Dim combobox As ComboBox = New ComboBox() From {
            New Item(Sub(ib) ib.WithLabel("Item1"))
        }

        Assert.That.ValueChangedIsRaisedOnce(combobox, Sub(cb) cb.First().Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub EventCanBeSubscribedTo()
        Dim combobox As ComboBox = New ComboBox() From {
            Item.Blank()
        }
        Dim before As EventHandler(Of BeforeTextChangeEventArgs) = Sub(s, e) Return
        Dim onChange As EventHandler(Of TextChangeEventArgs) = Sub(s, e) Return

        AddHandler combobox.BeforeTextChange, before
        AddHandler combobox.TextChanged, onChange

        RemoveHandler combobox.BeforeTextChange, before
        RemoveHandler combobox.TextChanged, onChange
    End Sub

    Public Shared Function BuildReadonlyComboBox() As ComboBox
        Return New ComboBox(Sub(cbb) cbb.
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

    Public Shared Function BuildReadonlyComboBoxII() As ComboBox
        Return New ComboBox(Sub(cbb) cbb.
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

    Public Shared Function BuildComboBox() As ComboBox
        Return New ComboBox(Sub(cbb) cbb.
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

    Public Shared Function BuildComboBoxII() As ComboBox
        Return New ComboBox(Sub(cbb) cbb.
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
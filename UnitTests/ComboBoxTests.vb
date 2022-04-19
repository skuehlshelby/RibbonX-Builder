Imports System.Drawing
Imports RibbonFactory
Imports RibbonFactory.Controls
Imports RibbonFactory.Controls
Imports RibbonFactory.Enums.ImageMSO

<TestClass()>
Public Class ComboBoxTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim combobox As ComboBox = New ComboBox(Sub(cbb) Return, template:=Nothing)

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim combobox As ComboBox = New ComboBox(configuration:=Nothing, template:=Nothing)

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Dim ribbon As Ribbon = CreateSimpleRibbon(
            New ComboBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           AsWideAs("AAAAAA").
                           WithImage(New Bitmap(48, 48), AddressOf GetImage).
                           WithText("The Text", AddressOf GetText, AddressOf OnChange).
                           ShowItemImages().
                           GetItemImageFrom(AddressOf GetItemImage).
                           GetItemIdFrom(AddressOf GetItemID).
                           GetItemLabelFrom(AddressOf GetItemLabel).
                           GetItemScreenTipFrom(AddressOf GetItemScreenTip).
                           GetItemSuperTipFrom(AddressOf GetItemSuperTip).
                           GetItemCountFrom(AddressOf GetItemCount)))

        Assert.IsTrue(ribbon.GetErrors().None)

        Debug.WriteLine(ribbon.RibbonX)
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Dim combobox As ComboBox = New ComboBox()

        ContainsNoNullValuesByDefaultHelper(combobox)

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim combobox As ComboBox =
            New ComboBox(Sub(cbb) cbb.
                           Visible().
                           Enabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           WithMaximumInputLength(40).
                           WithImage(Common.SadFace).
                           WithKeyTip("K2").
                           WithText("The Text", AddressOf GetText, AddressOf OnChange).
                           ShowItemImages())

        Assert.IsTrue(combobox.Enabled)
        Assert.IsTrue(combobox.Visible)
        Assert.IsTrue(combobox.ShowLabel)
        Assert.IsTrue(combobox.ShowImage)
        Assert.IsFalse(combobox.IsReadOnly)
        Assert.AreEqual(combobox.Label, "The Label")
        Assert.AreEqual(combobox.ScreenTip, "The Label")
        Assert.AreEqual(combobox.SuperTip, "The Supertip")
        Assert.AreEqual(combobox.MaxLength, CByte(40))
        Assert.AreEqual(combobox.Image, Common.SadFace)
        Assert.AreEqual(combobox.Text, "The Text")
        Assert.AreEqual(combobox.KeyTip, "K2")

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim combobox As ComboBox =
            New ComboBox(Sub(cbb) cbb.
                           Invisible().
                           Disabled().
                           WithLabel("The Label").
                           WithSuperTip("The Supertip").
                           HideLabel().
                           HideImage().
                           HideItemImages().
                           WithKeyTip("K2").
                           WithMaximumInputLength(40).
                           WithImage(Common.SadFace))

        Assert.ThrowsException(Of Exception)(Sub() combobox.Enabled = True)
        Assert.ThrowsException(Of Exception)(Sub() combobox.Visible = True)
        Assert.ThrowsException(Of Exception)(Sub() combobox.ShowLabel = True)
        Assert.ThrowsException(Of Exception)(Sub() combobox.ShowImage = True)
        Assert.ThrowsException(Of Exception)(Sub() combobox.Label = "New Label")
        Assert.ThrowsException(Of Exception)(Sub() combobox.ScreenTip = "New ScreenTip")
        Assert.ThrowsException(Of Exception)(Sub() combobox.SuperTip = "New SuperTip")
        Assert.ThrowsException(Of Exception)(Sub() combobox.Image = RibbonImage.Create(Common.FileFind))
        Assert.ThrowsException(Of Exception)(Sub() combobox.KeyTip = "5")

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim combobox As ComboBox = New ComboBox(Sub(cbb) cbb.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           WithLabel("The Label", AddressOf GetLabel).
                           WithSuperTip("The Supertip", AddressOf GetSuperTip).
                           WithKeyTip("K2", AddressOf GetKeyTip).
                           WithText("The Text", AddressOf GetText, AddressOf OnChange))

        Assert.IsTrue(combobox.Visible)
        combobox.Visible = False
        Assert.IsFalse(combobox.Visible)

        Assert.IsTrue(combobox.Enabled)
        combobox.Enabled = False
        Assert.IsFalse(combobox.Enabled)

        Assert.AreEqual(combobox.Label, "The Label")
        combobox.Label = "New Label"
        Assert.AreEqual(combobox.Label, "New Label")

        Assert.AreEqual(combobox.ScreenTip, "The Label")
        combobox.ScreenTip = "New ScreenTip"
        Assert.AreEqual(combobox.ScreenTip, "New ScreenTip")

        Assert.AreEqual(combobox.SuperTip, "The Supertip")
        combobox.SuperTip = "New SuperTip"
        Assert.AreEqual(combobox.SuperTip, "New SuperTip")

        Assert.AreEqual(combobox.KeyTip, "K2")
        combobox.KeyTip = "A"
        Assert.AreEqual(combobox.KeyTip, "A")

        Assert.AreEqual(combobox.Text, "The Text")
        combobox.Text = "The New Text"
        Assert.AreEqual(combobox.Text, "The New Text")

        Debug.WriteLine(combobox)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim firstCombobox As ComboBox = New ComboBox(Sub(cbb) cbb.WithImage(Common.Refresh).WithLabel("The Label").WithId("MyID"))
        Dim secondCombobox As ComboBox = New ComboBox(Nothing, template:=firstCombobox)

        Assert.AreEqual(secondCombobox.Image, Common.Refresh)
        Assert.AreEqual(secondCombobox.Label, "The Label")
        Assert.AreNotEqual(secondCombobox.ID, firstCombobox.ID)

        Debug.WriteLine(firstCombobox)
        Debug.WriteLine(secondCombobox)
    End Sub

    <TestMethod>
    Public Sub CanAddItems()
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

        Assert.That.ValueChangedIsRaised(combobox, Sub(cb) cb.Label = "New Label")
    End Sub

    <TestMethod>
    Public Sub ItemAddTriggersRefresh()
        Dim combobox As ComboBox = New ComboBox()

        Assert.That.ValueChangedIsRaised(combobox, Sub(cb) cb.Add(Item.Blank()))
    End Sub

    <TestMethod>
    Public Sub ChangingChildItemTriggersRefresh()
        Dim combobox As ComboBox = New ComboBox() From {
            Item.Blank()
        }

        Assert.That.ValueChangedIsRaised(combobox, Sub(cb) cb.First().Label = "New Label")
    End Sub

End Class
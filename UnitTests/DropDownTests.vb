Imports RibbonFactory
Imports RibbonFactory.BuilderInterfaces.API
Imports RibbonFactory.Controls

<TestClass()>
Public Class DropDownTests
    Inherits TestBase

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim dropdown As DropDown = New DropDown()
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim dropdown As DropDown = New DropDown(Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Dim ribbon As Ribbon = CreateSimpleRibbon(
            New DropDown(Sub(bb) bb.
                           Visible().
                           Enabled().
                           ShowImage().
                           ShowLabel().
                           WithLabel("Button").
                           WithSuperTip("Super").
                           WithImage(Enums.ImageMSO.Common.DollarSign)) From {Item.Blank})

        Assert.IsTrue(ribbon.GetErrors().None)

        Debug.WriteLine(ribbon.RibbonX)
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        ContainsNoNullValuesByDefaultHelper(New DropDown())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim dropDown As DropDown = New DropDown(Sub(bb As IDropdownBuilder) bb.
                           Visible().
                           Enabled().
                           ShowImage().
                           ShowLabel().
                           WithLabel("DropDown").
                           WithSuperTip("Super").
                           WithKeyTip("K2").
                           WithImage(Enums.ImageMSO.Common.DollarSign))

        Assert.IsTrue(dropDown.Enabled)
        Assert.IsTrue(dropDown.Visible)
        Assert.IsTrue(dropDown.ShowImage)
        Assert.IsTrue(dropDown.ShowLabel)
        Assert.AreEqual(dropDown.Label, "DropDown")
        Assert.AreEqual(dropDown.ScreenTip, dropDown.Label)
        Assert.AreEqual(dropDown.SuperTip, "Super")
        Assert.AreEqual(dropDown.KeyTip, "K2")
        Assert.AreEqual(dropDown.Image, Enums.ImageMSO.Common.DollarSign)

        Debug.WriteLine(dropDown)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Dim dropDown As DropDown = New DropDown(Sub(bb As IDropdownBuilder) bb.
                           Invisible().
                           Disabled().
                           WithLabel("DropDown").
                           WithSuperTip("Super").
                           WithKeyTip("K2").
                           WithImage(Enums.ImageMSO.Common.DollarSign).
                           HideImage().
                           HideLabel())

        Assert.ThrowsException(Of Exception)(Sub() dropDown.Visible = True)
        Assert.ThrowsException(Of Exception)(Sub() dropDown.Enabled = True)
        Assert.ThrowsException(Of Exception)(Sub() dropDown.ShowLabel = True)
        Assert.ThrowsException(Of Exception)(Sub() dropDown.ShowImage = True)
        Assert.ThrowsException(Of Exception)(Sub() dropDown.Label = "New Label")
        Assert.ThrowsException(Of Exception)(Sub() dropDown.ScreenTip = "New Screentip")
        Assert.ThrowsException(Of Exception)(Sub() dropDown.SuperTip = "New Supertip")
        Assert.ThrowsException(Of Exception)(Sub() dropDown.KeyTip = "K5")
        Assert.ThrowsException(Of Exception)(Sub() dropDown.Image = RibbonImage.Create(Enums.ImageMSO.Common.SadFace))

        Debug.WriteLine(dropDown)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim dropDown As DropDown = New DropDown(Sub(bb As IDropdownBuilder) bb.
                           Visible(AddressOf GetVisible).
                           Enabled(AddressOf GetEnabled).
                           ShowImage(AddressOf GetShowImage).
                           WithLabel("DropDown", AddressOf GetLabel).
                           ShowLabel(AddressOf GetShowLabel).
                           WithSuperTip("Super", AddressOf GetSuperTip).
                           WithKeyTip("K2", AddressOf GetKeyTip))

        dropDown.Enabled = False
        dropDown.Visible = False
        dropDown.ShowImage = False
        dropDown.ShowLabel = False
        dropDown.Label = "New Label"
        dropDown.ScreenTip = "New Label"
        dropDown.SuperTip = "New Label"
        dropDown.KeyTip = "K5"

        Debug.WriteLine(dropDown)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim dropDown As DropDown = New DropDown(Sub(bb As IDropdownBuilder) bb.
                           Invisible().
                           Disabled().
                           WithLabel("DropDown").
                           WithSuperTip("Super").
                           WithKeyTip("K2").
                           WithImage(Enums.ImageMSO.Common.DollarSign).
                           HideImage().
                           HideLabel())

        Dim anotherDropDown As DropDown = New DropDown(Nothing, dropDown)

        Assert.AreEqual(dropDown.Label, anotherDropDown.Label)
    End Sub

    <TestMethod>
    Public Sub CanAddItemsAfterDropdownCreation()
        Dim dropDown As DropDown = New DropDown() From {
            Item.Blank,
            Item.Blank,
            Item.Blank,
            Item.Blank
        }

        Assert.AreEqual(dropDown.Count, 4)
    End Sub

    <TestMethod>
    Public Sub CanClearItems()
        Dim items As Item() = {
            New Item(Sub(ib) ib.WithLabel("Item1")),
            New Item(Sub(ib) ib.WithLabel("Item2")),
            New Item(Sub(ib) ib.WithLabel("Item3")),
            New Item(Sub(ib) ib.WithLabel("Item4"))
        }

        Dim dropDown As DropDown = New DropDown()

        Array.ForEach(items, Sub(item) dropDown.Add(item))

        Assert.AreEqual(dropDown.Count, 4)

        Assert.AreEqual(dropDown.Selected, items.First())

        dropDown.Clear()

        Assert.AreEqual(dropDown.Count, 0)

        Assert.AreEqual(dropDown.Selected, Item.Blank)
    End Sub

End Class
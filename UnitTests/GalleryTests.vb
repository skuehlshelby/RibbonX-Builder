Imports RibbonX.Controls
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.SimpleTypes

<TestClass>
Public Class GalleryTests
    Inherits TestBase

    Private Const LABEL As String = "My Gallery"
    Private Const SCREENTIP As String = "My Gallery Screentip"
    Private Const SUPERTIP As String = "More Info About My Gallery"
    Private Const DESCRIPTION As String = "Even More Info About My Gallery"
    Private Const KEYTIP As String = "D"
    Private Const ITEM_HEIGHT As Integer = 48
    Private Const ITEM_WIDTH As Integer = 48

    <TestMethod>
    Public Overrides Sub NullTemplate_NoThrow()
        Dim control As Gallery = New Gallery(template:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub NullConfiguration_NoThrow()
        Dim control As Gallery = New Gallery(config:=Nothing)
    End Sub

    <TestMethod>
    Public Overrides Sub ProducesLegalRibbonX()
        Assert.That.ValidRibbonXIsProduced(BuildGallery())
    End Sub

    <TestMethod>
    Public Overrides Sub ContainsNoNullValuesByDefault()
        Assert.That.NoPropertiesAreNull(New Gallery())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesAreMappedCorrectly()
        Dim control As Gallery = BuildReadonlyGallery()

        Assert.IsTrue(control.Enabled)
        Assert.IsTrue(control.Visible)
        Assert.IsTrue(control.ShowLabel)
        Assert.IsTrue(control.ShowImage)
        Assert.AreEqual(control.Label, LABEL)
        Assert.AreEqual(control.ScreenTip, SCREENTIP)
        Assert.AreEqual(control.SuperTip, SUPERTIP)
        Assert.AreEqual(control.Description, DESCRIPTION)
        Assert.AreEqual(control.KeyTip, KEYTIP)
        Assert.AreEqual(control.Image, "img.png")
        Assert.AreEqual(control.ItemHeight, ITEM_HEIGHT)
        Assert.AreEqual(control.ItemWidth, ITEM_WIDTH)
        Assert.AreEqual(control.Size, ControlSize.Large)
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithoutCallbacksCannotBeModified()
        Assert.That.PropertiesCannotBeModified(BuildReadonlyGalleryII())
    End Sub

    <TestMethod>
    Public Overrides Sub PropertiesWithCallbacksCanBeModified()
        Dim control As Gallery = BuildGalleryII()

        Assert.That.PropertyMayBeModified(control, Function(c) c.Enabled, Not control.Enabled)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Visible, Not control.Visible)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ShowLabel, Not control.ShowLabel)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ShowImage, Not control.ShowImage)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Label, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.ScreenTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.SuperTip, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.Description, String.Empty)
        Assert.That.PropertyMayBeModified(control, Function(c) c.KeyTip, CType("G", KeyTip))
        Assert.That.PropertyMayBeModified(control, Function(c) c.Size, ControlSize.Large)
    End Sub

    <TestMethod>
    Public Overrides Sub TemplatePropertiesAreCopiedToNewControl()
        Dim control As Gallery = BuildReadonlyGallery()

        Assert.That.SharedPropertiesAreEqual(control, New DropDown(template:=control))
    End Sub

    Public Shared Function BuildReadonlyGallery() As Gallery
        Return New Gallery(config:=Sub(b) b.
                                       InsertAfter(Word.About).
                                       Enabled().
                                       Visible().
                                       WithLabel(LABEL).
                                       WithScreenTip(SCREENTIP).
                                       WithSuperTip(SUPERTIP).
                                       WithDescription(DESCRIPTION).
                                       WithKeyTip(KEYTIP).
                                       ShowLabel().
                                       WithImage("img.png").
                                       ShowImage().
                                       ShowLabelsOnChildItems().
                                       ShowItemImages().
                                       Large().
                                       WithItemDimensions(ITEM_HEIGHT, ITEM_WIDTH).
                                       WithRowCount(2).
                                       WithColumnCount(2),
                           buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()}) From {
                            Item.Blank, Item.Blank
                           }
    End Function

    Public Shared Function BuildReadonlyGalleryII() As Gallery
        Return New Gallery(config:=Sub(b) b.
                               InsertAfter(Word.About).
                               Disabled().
                               Invisible().
                               InvalidateContentOnDrop().
                               WithLabel(LABEL).
                               WithScreenTip(SCREENTIP).
                               WithSuperTip(SUPERTIP).
                               WithDescription(DESCRIPTION).
                               WithKeyTip(KEYTIP).
                               HideLabel().
                               WithImage("img.png").
                               HideImage().
                               HideLabelsOnChildItems().
                               HideItemImages().
                               Normal().
                               WithItemDimensions(ITEM_HEIGHT, ITEM_WIDTH).
                               WithRowCount(2).
                               WithColumnCount(2),
                   buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()}) From {
                    Item.Blank, Item.Blank
                   }
    End Function

    Public Shared Function BuildGallery() As Gallery
        Return New Gallery(config:=Sub(b) b.
                               InsertAfter(Word.About).
                               Enabled(AddressOf GetEnabledShared).
                               Visible(AddressOf GetVisibleShared).
                               WithLabel(LABEL, AddressOf GetLabelShared).
                               WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                               WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                               WithDescription(DESCRIPTION, AddressOf GetDescriptionShared).
                               WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                               ShowLabel(AddressOf GetShowLabelShared).
                               WithImage(BlankBitmap(), AddressOf GetImageShared).
                               ShowImage(AddressOf GetShowImageShared).
                               ShowLabelsOnChildItems().
                               ShowItemImages().
                               Large(AddressOf GetSizeShared).
                               WithItemDimensions(ITEM_HEIGHT, ITEM_WIDTH).
                               WithRowCount(2).
                               WithColumnCount(2),
                   buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()}) From {
                    Item.Blank, Item.Blank
                   }
    End Function

    Public Shared Function BuildGalleryII() As Gallery
        Return New Gallery(config:=Sub(b) b.
                       InsertAfter(Word.About).
                       Disabled(AddressOf GetEnabledShared).
                       Invisible(AddressOf GetVisibleShared).
                       WithLabel(LABEL, AddressOf GetLabelShared).
                       WithScreenTip(SCREENTIP, AddressOf GetScreenTipShared).
                       WithSuperTip(SUPERTIP, AddressOf GetSuperTipShared).
                       WithDescription(DESCRIPTION, AddressOf GetDescriptionShared).
                       WithKeyTip(KEYTIP, AddressOf GetKeyTipShared).
                       HideLabel(AddressOf GetShowLabelShared).
                       WithImage(BlankBitmap(), AddressOf GetImageShared).
                       HideImage(AddressOf GetShowImageShared).
                       ShowLabelsOnChildItems().
                       ShowItemImages().
                       Normal(AddressOf GetSizeShared).
                       WithItemDimensions(ITEM_HEIGHT, ITEM_WIDTH).
                       WithRowCount(2).
                       WithColumnCount(2),
           buttons:={ButtonTests.BuildReadonlyButton(), ButtonTests.BuildReadonlyButtonII()}) From {
            Item.Blank, Item.Blank
           }
    End Function
End Class

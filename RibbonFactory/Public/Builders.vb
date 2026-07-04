
Imports System.Drawing
Imports RibbonX.Api.Internal
Imports RibbonX.Callbacks
Imports RibbonX.ComTypes.StdOle
Imports RibbonX.Controls.BuiltIn
Imports RibbonX.Images.BuiltIn
Imports RibbonX.SimpleTypes

Namespace Api

    Namespace Internal

        Public Interface ITag(Of Out TBuilder)
            Function WithTag(tag As Object) As TBuilder
        End Interface

        Public Interface ITemplatable(Of Out TBuilder)
            Function FromTemplate(template As IRibbonElement) As TBuilder
        End Interface

        Public Interface IAddButton(Of Out TBuilder)
            Function Add(button As IButton) As TBuilder
        End Interface

        Public Interface IBoxStyle(Of Out TBuilder)
            Function Horizontal() As TBuilder
            Function Vertical() As TBuilder
        End Interface

        Public Interface IDescription(Of Out TBuilder)
            Function WithDescription(description As String) As TBuilder
            Function WithDescription(description As String, callback As FromControl(Of String)) As TBuilder
        End Interface

        Public Interface IEnabled(Of Out TBuilder)
            Function Enabled() As TBuilder
            Function Enabled(callback As FromControl(Of Boolean)) As TBuilder
            Function Disabled() As TBuilder
            Function Disabled(callback As FromControl(Of Boolean)) As TBuilder
        End Interface

        Public Interface IGetItemCount(Of Out TBuilder)
            Function GetItemCountFrom(callback As FromControl(Of Integer)) As TBuilder
        End Interface

        Public Interface IGetItemId(Of Out TBuilder)
            Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As TBuilder
        End Interface

        Public Interface IGetItemImage(Of Out TBuilder)
            Function GetItemImageFrom(callback As FromControlAndIndex(Of IPictureDisp)) As TBuilder
            Function GetItemImageFrom(callback As FromControlAndIndex(Of String)) As TBuilder
        End Interface

        Public Interface IGetItemLabel(Of Out TBuilder)
            Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As TBuilder
        End Interface

        Public Interface IGetItemScreentip(Of Out TBuilder)
            Function GetItemScreenTipFrom(callback As FromControlAndIndex(Of String)) As TBuilder
        End Interface

        Public Interface IGetItemSupertip(Of Out TBuilder)
            Function GetItemSuperTipFrom(callback As FromControlAndIndex(Of String)) As TBuilder
        End Interface

        Public Interface IGetSelected(Of Out TBuilder)
            Function GetSelectedItemIdFrom(getSelected As FromControl(Of String), setSelected As SelectionChanged) As TBuilder
            Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As TBuilder

        End Interface

        Public Interface IID(Of TBuilder)
            Function WithId(id As String) As TBuilder
            Function WithIdQ([namespace] As String, id As String) As TBuilder
            Function WithIdMso(mso As MSO) As TBuilder
        End Interface

        Public Interface IImage(Of Out TBuilder)
            Function WithImage(image As ImageMSO) As TBuilder
            Function WithImage(image As ImageMSO, callback As FromControl(Of String)) As TBuilder
            Function WithImage(imageId As String, image As Bitmap) As TBuilder
            Function WithImage(imageId As String, image As Icon) As TBuilder
            Function WithImage(image As Bitmap, callback As FromControl(Of IPictureDisp)) As TBuilder
            Function WithImage(image As Icon, callback As FromControl(Of IPictureDisp)) As TBuilder
        End Interface

        Public Interface IInsert(Of Out TBuilder)
            Function InsertBefore(builtInControl As MSO) As TBuilder
            Function InsertBefore(qualifiedControl As IRibbonElement) As TBuilder
            Function InsertAfter(builtInControl As MSO) As TBuilder
            Function InsertAfter(qualifiedControl As IRibbonElement) As TBuilder
        End Interface

        Public Interface IInvalidateContentOnDrop(Of Out TBuilder)
            Function InvalidateContentOnDrop() As TBuilder
        End Interface

        Public Interface IItemDimensions(Of Out TBuilder)
            Function WithItemDimensions(size As Size) As TBuilder
            Function WithItemDimensions(height As Integer, width As Integer) As TBuilder
            Function WithItemDimensions(size As Size, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As TBuilder
            Function WithItemDimensions(height As Integer, width As Integer, heightCallback As FromControl(Of Integer), widthCallback As FromControl(Of Integer)) As TBuilder
        End Interface

        Public Interface IItemSize(Of Out TBuilder)
            Function WithLargeItems() As TBuilder
            Function WithNormalSizeItems() As TBuilder
        End Interface

        Public Interface IKeyTip(Of Out TBuilder)
            Function WithKeyTip(keyTip As KeyTip) As TBuilder
            Function WithKeyTip(keyTip As KeyTip, callback As FromControl(Of KeyTip)) As TBuilder
        End Interface

        Public Interface ILabelOnly(Of Out TBuilder)
            Function WithLabel(label As String) As TBuilder
        End Interface

        Public Interface ILabelScreenTipSuperTip(Of Out TBuilder)
            Function WithLabel(label As String, Optional copyToScreenTip As Boolean = True) As TBuilder
            Function WithLabel(label As String, callback As FromControl(Of String), Optional copyToScreenTip As Boolean = True) As TBuilder
            Function WithScreenTip(screenTip As String) As TBuilder
            Function WithScreenTip(screenTip As String, callback As FromControl(Of String)) As TBuilder
            Function WithSuperTip(superTip As String) As TBuilder
            Function WithSuperTip(superTip As String, callback As FromControl(Of String)) As TBuilder
        End Interface

        Public Interface ILoadImage(Of Out TBuilder)
            Function LoadImagesFrom(callback As LoadImage) As TBuilder
        End Interface

        Public Interface IMaxLength(Of Out TBuilder)
            Function WithMaximumInputLength(maxLength As Byte) As TBuilder
        End Interface

        Public Interface IOnActionClick(Of TElement As IRibbonElement, Out TBuilder)
            Function ThatDoes(action As Action(Of TElement), callback As OnAction) As TBuilder
        End Interface

        Public Interface IOnActionSelectionChanged(Of Out TBuilder)
            Function ThatDoes(action As Action, callback As SelectionChanged) As TBuilder
        End Interface

        Public Interface IOnActionToggle(Of Out TBuilder, TRibbonElement As IRibbonElement)
            Function ThatDoes(action As Action(Of TRibbonElement), callback As ButtonPressed) As TBuilder
        End Interface

        Public Interface IOnChange(Of TRibbonElement As IRibbonElement, Out TBuilder)
            Function ThatDoes(action As Action(Of TRibbonElement), callback As TextChanged) As TBuilder
        End Interface

        Public Interface IOnLoad(Of Out TBuilder)
            Function OnLoad(callback As OnLoad) As TBuilder
        End Interface

        Public Interface IRowsAndColumns(Of Out TBuilder)
            Function WithRowCount(rowCount As Integer) As TBuilder
            Function WithColumnCount(columnCount As Integer) As TBuilder
        End Interface

        Public Interface IShowImage(Of Out TBuilder)
            Function ShowImage() As TBuilder
            Function ShowImage(getShowImage As FromControl(Of Boolean)) As TBuilder
            Function HideImage() As TBuilder
            Function HideImage(getShowImage As FromControl(Of Boolean)) As TBuilder
        End Interface

        Public Interface IShowInRibbon(Of Out TBuilder)
            Function DoNotShowInRibbon() As TBuilder
        End Interface

        Public Interface IShowItemImage(Of Out TBuilder)
            Function ShowItemImages() As TBuilder
            Function HideItemImages() As TBuilder
        End Interface

        Public Interface IShowItemLabel(Of Out TBuilder)
            Function ShowLabelsOnChildItems() As TBuilder
            Function HideLabelsOnChildItems() As TBuilder
        End Interface

        Public Interface IShowLabel(Of Out TBuilder)
            Function ShowLabel() As TBuilder
            Function ShowLabel(callback As FromControl(Of Boolean)) As TBuilder
            Function HideLabel() As TBuilder
            Function HideLabel(callback As FromControl(Of Boolean)) As TBuilder
        End Interface

        Public Interface ISize(Of Out TBuilder)
            Function Large() As TBuilder
            Function Large(callback As FromControl(Of ControlSize)) As TBuilder
            Function Normal() As TBuilder
            Function Normal(callback As FromControl(Of ControlSize)) As TBuilder
        End Interface

        Public Interface ISizeString(Of Out TBuilder)
            Function AsWideAs(sizeString As String) As TBuilder
        End Interface

        Public Interface IStartFromScratch(Of Out TBuilder)
            Function StartFromScratch() As TBuilder
        End Interface

        Public Interface IVisible(Of Out TBuilder)
            Function Visible() As TBuilder
            Function Visible(callback As FromControl(Of Boolean)) As TBuilder
            Function Invisible() As TBuilder
            Function Invisible(callback As FromControl(Of Boolean)) As TBuilder
        End Interface

        Public Interface IActionBuilder(Of TControl As IRibbonElement)
            Function [Do](ParamArray action As Action()) As IActionBuilder(Of TControl)
            Function [Do](ParamArray action As Action(Of TControl)()) As IActionBuilder(Of TControl)
            Function ButFirst(ParamArray action As Action()) As IActionBuilder(Of TControl)
            Function ButFirst(ParamArray action As Func(Of Boolean)()) As IActionBuilder(Of TControl)
            Function ButFirst(ParamArray action As Func(Of TControl, Boolean)()) As IActionBuilder(Of TControl)
        End Interface

        Public Interface IActionBuilder(Of TControl As IRibbonElement, TData)
            Inherits IActionBuilder(Of TControl)
            Overloads Function [Do](ParamArray action As Action(Of TData)()) As IActionBuilder(Of TControl)
            Overloads Function [Do](ParamArray action As Action(Of TControl, TData)()) As IActionBuilder(Of TControl)
            Overloads Function ButFirst(ParamArray action As Func(Of TData, Boolean)()) As IActionBuilder(Of TControl)
            Overloads Function ButFirst(ParamArray action As Func(Of TControl, TData, Boolean)()) As IActionBuilder(Of TControl)
        End Interface

    End Namespace

    Public Interface IBoxBuilder
        Inherits ITag(Of IBoxBuilder)
        Inherits ITemplatable(Of IBoxBuilder)
        Inherits IBoxStyle(Of IBoxBuilder)
        Inherits IVisible(Of IBoxBuilder)
        Function WithControls(ParamArray controls() As IBoxAddable) As IBoxBuilder
    End Interface

    Public Interface IButtonBuilder
        Inherits IID(Of IButtonBuilder)
        Inherits ITag(Of IButtonBuilder)
        Inherits ITemplatable(Of IButtonBuilder)
        Inherits IEnabled(Of IButtonBuilder)
        Inherits IVisible(Of IButtonBuilder)
        Inherits IInsert(Of IButtonBuilder)
        Inherits ILabelScreenTipSuperTip(Of IButtonBuilder)
        Inherits IShowLabel(Of IButtonBuilder)
        Inherits IShowImage(Of IButtonBuilder)
        Inherits ISize(Of IButtonBuilder)
        Inherits IImage(Of IButtonBuilder)
        Inherits IDescription(Of IButtonBuilder)
        Inherits IKeyTip(Of IButtonBuilder)
        Function OnClick(callback As OnAction) As IButtonBuilder
        Function OnClick(callback As OnAction, action As Action(Of IActionBuilder(Of IButton))) As IButtonBuilder
    End Interface

    Public Interface IButtonGroupBuilder
        Inherits IID(Of IButtonGroupBuilder)
        Inherits ITag(Of IButtonGroupBuilder)
        Inherits ITemplatable(Of IButtonGroupBuilder)
        Inherits IInsert(Of IButtonGroupBuilder)
        Inherits IVisible(Of IButtonGroupBuilder)
        Function WithControls(ParamArray controls() As IButtonGroupAddable) As IButtonGroupBuilder
    End Interface

    Public Interface ICheckBoxBuilder
        Inherits IID(Of ICheckBoxBuilder)
        Inherits ITag(Of ICheckBoxBuilder)
        Inherits ITemplatable(Of ICheckBoxBuilder)
        Inherits IEnabled(Of ICheckBoxBuilder)
        Inherits IVisible(Of ICheckBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of ICheckBoxBuilder)
        Inherits IInsert(Of ICheckBoxBuilder)
        Inherits IKeyTip(Of ICheckBoxBuilder)
        Inherits IDescription(Of ICheckBoxBuilder)
        Function OnCheck(initialValue As Boolean, getChecked As FromControl(Of Boolean), setChecked As ButtonPressed) As ICheckBoxBuilder
        Function OnCheck(initialValue As Boolean, getChecked As FromControl(Of Boolean), setChecked As ButtonPressed, action As Action(Of IActionBuilder(Of ICheckbox, Boolean))) As ICheckBoxBuilder
    End Interface

    Public Interface IComboBoxBuilder
        Inherits IID(Of IComboBoxBuilder)
        Inherits ITag(Of IComboBoxBuilder)
        Inherits ITemplatable(Of IComboBoxBuilder)
        Inherits IEnabled(Of IComboBoxBuilder)
        Inherits IVisible(Of IComboBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of IComboBoxBuilder)
        Inherits IShowLabel(Of IComboBoxBuilder)
        Inherits IKeyTip(Of IComboBoxBuilder)
        Inherits IImage(Of IComboBoxBuilder)
        Inherits IShowImage(Of IComboBoxBuilder)
        Inherits IMaxLength(Of IComboBoxBuilder)
        Inherits ISizeString(Of IComboBoxBuilder)
        Inherits IInsert(Of IComboBoxBuilder)
        Inherits IGetItemId(Of IComboBoxBuilder)
        Inherits IGetItemCount(Of IComboBoxBuilder)
        Inherits IGetItemLabel(Of IComboBoxBuilder)
        Inherits IGetItemScreentip(Of IComboBoxBuilder)
        Inherits IGetItemSupertip(Of IComboBoxBuilder)
        Inherits IGetItemImage(Of IComboBoxBuilder)
        Inherits IShowItemImage(Of IComboBoxBuilder)
        Function WithText(initialValue As String, getText As FromControl(Of String), setText As TextChanged) As IComboBoxBuilder
        Function WithText(initialValue As String, getText As FromControl(Of String), setText As TextChanged, action As Action(Of IActionBuilder(Of IComboBox, String))) As IComboBoxBuilder
    End Interface

    Public Interface IDialogLauncherBuilder
        Inherits ITag(Of IDialogLauncherBuilder)
        Function WithButton(button As IButton) As IDialogLauncherBuilder
    End Interface

    Public Interface IDropdownBuilder
        Inherits IID(Of IDropdownBuilder)
        Inherits ITag(Of IDropdownBuilder)
        Inherits ITemplatable(Of IDropdownBuilder)
        Inherits IEnabled(Of IDropdownBuilder)
        Inherits IInsert(Of IDropdownBuilder)
        Inherits IVisible(Of IDropdownBuilder)
        Inherits ILabelScreenTipSuperTip(Of IDropdownBuilder)
        Inherits IKeyTip(Of IDropdownBuilder)
        Inherits IShowLabel(Of IDropdownBuilder)
        Inherits IImage(Of IDropdownBuilder)
        Inherits IShowImage(Of IDropdownBuilder)
        Inherits ISizeString(Of IDropdownBuilder)
        Inherits IGetItemId(Of IDropdownBuilder)
        Inherits IGetItemCount(Of IDropdownBuilder)
        Inherits IGetItemLabel(Of IDropdownBuilder)
        Inherits IShowItemLabel(Of IDropdownBuilder)
        Inherits IGetItemScreentip(Of IDropdownBuilder)
        Inherits IGetItemSupertip(Of IDropdownBuilder)
        Inherits IGetItemImage(Of IDropdownBuilder)
        Inherits IShowItemImage(Of IDropdownBuilder)
        Inherits IGetSelected(Of IDropdownBuilder)

        Function WithBlank(blank As IItem) As IDropdownBuilder
        Function WithButtons(ParamArray buttons As IButton()) As IDropdownBuilder
        Function OnSelectionChange(getSelected As FromControl(Of String), setSelected As SelectionChanged) As IDropdownBuilder
        Function OnSelectionChange(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As IDropdownBuilder
        Function OnSelectionChange(getSelected As FromControl(Of Integer), setSelected As SelectionChanged, action As Action(Of IActionBuilder(Of IDropDown, IItem))) As IDropdownBuilder
        Function OnSelectionChange(getSelected As FromControl(Of String), setSelected As SelectionChanged, action As Action(Of IActionBuilder(Of IDropDown, IItem))) As IDropdownBuilder
    End Interface

    Public Interface IEditBoxBuilder
        Inherits IID(Of IEditBoxBuilder)
        Inherits ITag(Of IEditBoxBuilder)
        Inherits ITemplatable(Of IEditBoxBuilder)
        Inherits IEnabled(Of IEditBoxBuilder)
        Inherits IVisible(Of IEditBoxBuilder)
        Inherits IInsert(Of IEditBoxBuilder)
        Inherits ILabelScreenTipSuperTip(Of IEditBoxBuilder)
        Inherits IKeyTip(Of IEditBoxBuilder)
        Inherits IImage(Of IEditBoxBuilder)
        Inherits IShowLabel(Of IEditBoxBuilder)
        Inherits IMaxLength(Of IEditBoxBuilder)
        Inherits ISizeString(Of IEditBoxBuilder)
        Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged) As IEditBoxBuilder
        Function WithText(text As String, getText As FromControl(Of String), setText As TextChanged, action As Action(Of IActionBuilder(Of IEditBox, String))) As IEditBoxBuilder
    End Interface

    Public Interface IGalleryBuilder
        Inherits IID(Of IGalleryBuilder)
        Inherits ITag(Of IGalleryBuilder)
        Inherits ITemplatable(Of IGalleryBuilder)
        Inherits IEnabled(Of IGalleryBuilder)
        Inherits IVisible(Of IGalleryBuilder)
        Inherits IInsert(Of IGalleryBuilder)
        Inherits ILabelScreenTipSuperTip(Of IGalleryBuilder)
        Inherits IShowLabel(Of IGalleryBuilder)
        Inherits IDescription(Of IGalleryBuilder)
        Inherits IShowInRibbon(Of IGalleryBuilder)
        Inherits IInvalidateContentOnDrop(Of IGalleryBuilder)
        Inherits IImage(Of IGalleryBuilder)
        Inherits IShowImage(Of IGalleryBuilder)
        Inherits IKeyTip(Of IGalleryBuilder)
        Inherits ISize(Of IGalleryBuilder)
        Inherits IItemDimensions(Of IGalleryBuilder)
        Inherits IRowsAndColumns(Of IGalleryBuilder)
        Inherits IGetItemId(Of IGalleryBuilder)
        Inherits IGetItemCount(Of IGalleryBuilder)
        Inherits IGetItemLabel(Of IGalleryBuilder)
        Inherits IShowItemLabel(Of IGalleryBuilder)
        Inherits IGetItemScreentip(Of IGalleryBuilder)
        Inherits IGetItemSupertip(Of IGalleryBuilder)
        Inherits IGetItemImage(Of IGalleryBuilder)
        Inherits IShowItemImage(Of IGalleryBuilder)
        Inherits IGetSelected(Of IGalleryBuilder)
        Function WithBlank(blank As IItem) As IGalleryBuilder
        Function WithButtons(ParamArray buttons As IButton()) As IGalleryBuilder
        Function OnSelectionChange(getSelected As FromControl(Of String), setSelected As SelectionChanged) As IGalleryBuilder
        Function OnSelectionChange(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As IGalleryBuilder
        Function OnSelectionChange(getSelected As FromControl(Of Integer), setSelected As SelectionChanged, action As Action(Of IActionBuilder(Of IGallery, IItem))) As IGalleryBuilder
        Function OnSelectionChange(getSelected As FromControl(Of String), setSelected As SelectionChanged, action As Action(Of IActionBuilder(Of IGallery, IItem))) As IGalleryBuilder
    End Interface

    Public Interface IGroupBuilder
        Inherits IID(Of IGroupBuilder)
        Inherits ITag(Of IGroupBuilder)
        Inherits ITemplatable(Of IGroupBuilder)
        Inherits IInsert(Of IGroupBuilder)
        Inherits IVisible(Of IGroupBuilder)
        Inherits ILabelScreenTipSuperTip(Of IGroupBuilder)
        Inherits IKeyTip(Of IGroupBuilder)
        Inherits IImage(Of IGroupBuilder)
        Function WithControls(ParamArray controls() As IGroupAddable) As IGroupBuilder
    End Interface

    Public Interface IItemBuilder
        Inherits ITag(Of IItemBuilder)
        Inherits ITemplatable(Of IItemBuilder)
        Function WithId(id As String) As IItemBuilder
        Function WithLabel(label As String, Optional copyToScreentip As Boolean = True) As IItemBuilder
        Function WithScreenTip(screenTip As String) As IItemBuilder
        Function WithSuperTip(superTip As String) As IItemBuilder
        Function WithImage(image As Bitmap) As IItemBuilder
        Function WithImage(image As Icon) As IItemBuilder
    End Interface

    Public Interface ILabelControlBuilder
        Inherits IID(Of ILabelControlBuilder)
        Inherits ITag(Of ILabelControlBuilder)
        Inherits ITemplatable(Of ILabelControlBuilder)
        Inherits IEnabled(Of ILabelControlBuilder)
        Inherits IInsert(Of ILabelControlBuilder)
        Inherits IVisible(Of ILabelControlBuilder)
        Inherits ILabelScreenTipSuperTip(Of ILabelControlBuilder)
        Inherits IShowLabel(Of ILabelControlBuilder)
    End Interface

    Public Interface IMenuBuilder
        Inherits ITag(Of IMenuBuilder)
        Inherits ITemplatable(Of IMenuBuilder)
        Inherits IInsert(Of IMenuBuilder)
        Inherits IEnabled(Of IMenuBuilder)
        Inherits IVisible(Of IMenuBuilder)
        Inherits ILabelScreenTipSuperTip(Of IMenuBuilder)
        Inherits IShowLabel(Of IMenuBuilder)
        Inherits IImage(Of IMenuBuilder)
        Inherits IShowImage(Of IMenuBuilder)
        Inherits IItemSize(Of IMenuBuilder)
        Inherits IDescription(Of IMenuBuilder)
        Inherits ISize(Of IMenuBuilder)
        Inherits IKeyTip(Of IMenuBuilder)
        Function WithControls(ParamArray controls() As IMenuAddable) As IMenuBuilder
    End Interface

    Public Interface IMenuSeparatorBuilder
        Inherits IID(Of IMenuSeparatorBuilder)
        Inherits ITag(Of IMenuSeparatorBuilder)
        Inherits ITemplatable(Of IMenuSeparatorBuilder)
        Inherits IInsert(Of IMenuSeparatorBuilder)
        Function WithTitle(title As String) As IMenuSeparatorBuilder
        Function WithTitle(title As String, callback As FromControl(Of String)) As IMenuSeparatorBuilder
    End Interface

    Public Interface IRibbonBuilder
        Inherits IStartFromScratch(Of IRibbonBuilder)
        Inherits IOnLoad(Of IRibbonBuilder)
        Inherits ILoadImage(Of IRibbonBuilder)
        Function WithTabs(ParamArray tabs As ITab()) As IRibbonBuilder
    End Interface

    Public Interface ISeparatorBuilder
        Inherits IID(Of ISeparatorBuilder)
        Inherits ITag(Of ISeparatorBuilder)
        Inherits ITemplatable(Of ISeparatorBuilder)
        Inherits IInsert(Of ISeparatorBuilder)
        Inherits IVisible(Of ISeparatorBuilder)
    End Interface

    Public Interface ISplitButtonBuilder
        Inherits IID(Of ISplitButtonBuilder)
        Inherits ITag(Of ISplitButtonBuilder)
        Inherits ITemplatable(Of ISplitButtonBuilder)
        Inherits IInsert(Of ISplitButtonBuilder)
        Inherits IVisible(Of ISplitButtonBuilder)
        Inherits IEnabled(Of ISplitButtonBuilder)
        Inherits IKeyTip(Of ISplitButtonBuilder)
        Inherits IShowLabel(Of ISplitButtonBuilder)
        Inherits ISize(Of ISplitButtonBuilder)
        Function WithButtonAndMenu(button As IButton, menu As IMenu) As ISplitButtonBuilder
        Function WithButtonAndMenu(button As IToggleButton, menu As IMenu) As ISplitButtonBuilder
    End Interface

    Public Interface ITabBuilder
        Inherits ITag(Of ITabBuilder)
        Inherits IInsert(Of ITabBuilder)
        Inherits ITemplatable(Of ITabBuilder)
        Inherits IKeyTip(Of ITabBuilder)
        Inherits ILabelOnly(Of ITabBuilder)
        Inherits IVisible(Of ITabBuilder)
        Function WithGroups(ParamArray groups As IGroup()) As ITabBuilder
    End Interface

    Public Interface IToggleButtonBuilder
        Inherits IID(Of IToggleButtonBuilder)
        Inherits ITag(Of IToggleButtonBuilder)
        Inherits ITemplatable(Of IToggleButtonBuilder)
        Inherits IVisible(Of IToggleButtonBuilder)
        Inherits IEnabled(Of IToggleButtonBuilder)
        Inherits IInsert(Of IToggleButtonBuilder)
        Inherits ILabelScreenTipSuperTip(Of IToggleButtonBuilder)
        Inherits IShowLabel(Of IToggleButtonBuilder)
        Inherits IKeyTip(Of IToggleButtonBuilder)
        Inherits IImage(Of IToggleButtonBuilder)
        Inherits IShowImage(Of IToggleButtonBuilder)
        Inherits IDescription(Of IToggleButtonBuilder)
        Inherits ISize(Of IToggleButtonBuilder)
        Function OnToggle(initialValue As Boolean, getToggled As FromControl(Of Boolean), setToggled As ButtonPressed, Optional action As Action(Of IActionBuilder(Of IToggleButton, Boolean)) = Nothing) As IToggleButtonBuilder
    End Interface

End Namespace
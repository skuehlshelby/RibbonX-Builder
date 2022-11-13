Imports System.ComponentModel
Imports RibbonX.Builders
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.Controls.Properties

Namespace Api

    Public Interface IRibbonFactory
        Function Box(Optional builder As Action(Of IBoxBuilder) = Nothing) As IBox
        Function Button(Optional builder As Action(Of IButtonBuilder) = Nothing) As IButton
        Function ButtonGroup(Optional builder As Action(Of IButtonGroupBuilder) = Nothing) As IButtonGroup
        Function CheckBox(Optional builder As Action(Of ICheckBoxBuilder) = Nothing) As ICheckbox
        Function ComboBox(Optional builder As Action(Of IComboBoxBuilder) = Nothing) As IComboBox
        Function DialogLauncher(button As IButton, Optional tag As Object = Nothing) As IDialogLauncher
        Function DropDown(Optional builder As Action(Of IDropdownBuilder) = Nothing) As IDropDown
        Function EditBox(Optional builder As Action(Of IEditBoxBuilder) = Nothing) As IEditBox
        Function Gallery(Optional builder As Action(Of IGalleryBuilder) = Nothing) As IGallery
        Function Group(Optional builder As Action(Of IGroupBuilder) = Nothing) As IGroup
        Function Item(Optional builder As Action(Of IItemBuilder) = Nothing) As IItem
        Function LabelControl(Optional builder As Action(Of ILabelControlBuilder) = Nothing) As ILabelControl
        Function Menu(Optional builder As Action(Of IMenuBuilder) = Nothing) As IMenu
        Function MenuSeparator(Optional builder As Action(Of IMenuSeparatorBuilder) = Nothing) As IMenuSeparator
        Function Ribbon(Optional builder As Action(Of IRibbonBuilder) = Nothing) As IRibbon
        Function Separator(Optional builder As Action(Of ISeparatorBuilder) = Nothing) As ISeparator
        Function SplitButton(Optional builder As Action(Of ISplitButtonBuilder) = Nothing) As ISplitButton
        Function Tab(Optional builder As Action(Of ITabBuilder) = Nothing) As ITab
        Function ToggleButton(Optional builder As Action(Of IToggleButtonBuilder) = Nothing) As IToggleButton
    End Interface

    Public Interface IRibbonElement
        Inherits IEquatable(Of IRibbonElement)
        Inherits INotifyPropertyChanged
        Inherits ICloneable
        ReadOnly Property Id As String
        Property Tag As Object
        Function ConvertToXml() As String
        Function GetChildren() As ICollection(Of IRibbonElement)
        Function SuspendRefreshing() As IDisposable
        Function Bind(binding As IBinding) As IDisposable
    End Interface

    Public Interface IBox
        Inherits IRibbonElement
        Inherits IReadOnlyCollection(Of IRibbonElement)
        Inherits IVisible
        Inherits IBoxStyle
    End Interface

    Public Interface IButton
        Inherits IRibbonElement
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IShowLabel
        Inherits IKeyTip
        Inherits IDescription
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IImage
        Inherits IShowImage
        Inherits ISize
        Inherits IClickable
    End Interface

    Public Interface IButtonGroup
        Inherits IRibbonElement
        Inherits IReadOnlyCollection(Of IRibbonElement)
        Inherits IVisible
    End Interface

    Public Interface ICheckbox
        Inherits IRibbonElement
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IDescription
        Inherits IKeyTip
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IChecked
        Event Checking As EventHandler
        Event Checked As EventHandler
    End Interface

    Public Interface IComboBox
        Inherits IRibbonElement
        Inherits ICollection(Of IItem)
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IShowLabel
        Inherits IKeyTip
        Inherits IImage
        Inherits IShowImage
        Inherits IMaxLength
        Inherits IText
        Event SelectionChanging As EventHandler
        Event SelectionChanged As EventHandler
    End Interface

    Public Interface IDialogLauncher
        Inherits IRibbonElement
    End Interface

    Public Interface IDropDown
        Inherits IRibbonElement
        Inherits ICollection(Of IItem)
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IShowLabel
        Inherits IKeyTip
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits ISelect
        Inherits IImage
        Inherits IShowImage
        Event SelectionChanging As EventHandler
        Event SelectionChanged As EventHandler
    End Interface

    Public Interface IEditBox
        Inherits IRibbonElement
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IShowLabel
        Inherits IKeyTip
        Inherits IImage
        Inherits IShowImage
        Inherits IText
        Inherits IMaxLength
        Event Changing As EventHandler
        Event Changed As EventHandler
    End Interface

    Public Interface IGallery
        Inherits IRibbonElement
        Inherits ICollection(Of IItem)
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IShowLabel
        Inherits IKeyTip
        Inherits IDescription
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IImage
        Inherits IShowImage
        Inherits ISelect
        Inherits ISize
        Inherits IItemDimensions
        Inherits IRowsAndColumns
        Event Clicking As EventHandler
        Event Clicked As EventHandler
    End Interface

    Public Interface IGroup
        Inherits IRibbonElement
        Inherits IReadOnlyCollection(Of IRibbonElement)
        Inherits ILabel
        Inherits IVisible
        Inherits IKeyTip
        Inherits IImage
        Inherits IScreenTip
        Inherits ISuperTip
    End Interface

    Public Interface IItem
        Inherits IRibbonElement
        Inherits ILabel
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IImage
    End Interface

    Public Interface ILabelControl
        Inherits IRibbonElement
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IShowLabel
        Inherits IScreenTip
        Inherits ISuperTip
    End Interface

    Public Interface IMenu
        Inherits IRibbonElement
        Inherits IReadOnlyCollection(Of IRibbonElement)
        Inherits IEnabled
        Inherits IVisible
        Inherits ILabel
        Inherits IShowLabel
        Inherits IDescription
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IKeyTip
        Inherits IImage
        Inherits IShowImage
        Inherits ISize
        Inherits IItemSize
    End Interface

    Public Interface IMenuSeparator
        Inherits IRibbonElement
        Inherits ITitle
    End Interface

    Public Interface IRibbon
        Inherits IRibbonUI
        Inherits IReadOnlyCollection(Of IRibbonElement)
    End Interface

    Public Interface ISeparator
        Inherits IRibbonElement
        Inherits IVisible
    End Interface

    Public Interface ISplitButton
        Inherits IRibbonElement
        Inherits IReadOnlyCollection(Of IRibbonElement)
        Inherits IEnabled
        Inherits IVisible
        Inherits IKeyTip
        Inherits ISize
        Inherits IShowLabel
    End Interface

    Public Interface ITab
        Inherits IRibbonElement
        Inherits ICollection(Of IGroup)
        Inherits IVisible
        Inherits IKeyTip
        Inherits ILabel
    End Interface

    Public Interface IToggleButton
        Inherits IRibbonElement
        Inherits IEnabled
        Inherits IVisible
        Inherits IScreenTip
        Inherits ISuperTip
        Inherits IDescription
        Inherits ILabel
        Inherits IShowLabel
        Inherits IImage
        Inherits IShowImage
        Inherits ISize
        Inherits IChecked
        Inherits IKeyTip
    End Interface

End Namespace
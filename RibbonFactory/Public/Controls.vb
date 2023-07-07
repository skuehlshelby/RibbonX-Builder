Imports System.ComponentModel
Imports System.Linq.Expressions
Imports RibbonX.ComTypes.Microsoft.Office.Core
Imports RibbonX.Properties
Imports RibbonX.Testing

Public Enum ExcludedAttributes
    None
    Size
    SizeAndVisibility
End Enum

Public Interface IRibbonElement
    Inherits IEquatable(Of IRibbonElement)
    Inherits INotifyPropertyChanged
    Inherits ICloneable
    Event RefreshNeeded As EventHandler(Of RefreshNeededEventArgs)
    ReadOnly Property Id As String
    Property Tag As Object
    Function ToXml(Optional excluded As ExcludedAttributes = ExcludedAttributes.None) As String
    Function SuspendRefreshing() As IDisposable
End Interface

Public Interface IBoxAddable
    Inherits IRibbonElement
End Interface

Public Interface IBox
    Inherits IRibbonElement
    Inherits IReadOnlyCollection(Of IRibbonElement)
    Inherits IVisible
    Inherits IBoxStyle
    Inherits IBoxAddable
    Inherits IGroupAddable
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
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IMenuAddable
    Inherits IGroupAddable
    Sub Bind(Of TTarget As Class)(target As TTarget, ParamArray bindings As Expression(Of Action(Of IButton, TTarget))())
    Event Clicking As EventHandler(Of CancelableEventArgs)
    Event Clicked As EventHandler
End Interface

Public Interface IButtonGroupAddable
    Inherits IRibbonElement
End Interface

Public Interface IButtonGroup
    Inherits IRibbonElement
    Inherits IReadOnlyCollection(Of IRibbonElement)
    Inherits IVisible
    Inherits IBoxAddable
    Inherits IGroupAddable
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
    Inherits IBoxAddable
    Inherits IMenuAddable
    Inherits IGroupAddable
    Event Checking As EventHandler(Of CancelableEventArgs(Of Boolean))
    Event Checked As EventHandler(Of EventArgs(Of Boolean))
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
    Inherits IItemTemplateable
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IGroupAddable
    Event Changing As EventHandler(Of CancelableEventArgs(Of String))
    Event Changed As EventHandler(Of EventArgs(Of String))
End Interface

Public Interface IDialogLauncher
    Inherits IRibbonElement
    Inherits IGroupAddable
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
    Inherits IItemTemplateable
    Inherits IBoxAddable
    Inherits IGroupAddable
    ReadOnly Property Buttons As ICollection(Of IButton)
    Event SelectionChanging As EventHandler(Of CancelableEventArgs(Of IItem))
    Event SelectionChanged As EventHandler(Of EventArgs(Of IItem))
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
    Inherits IBoxAddable
    Inherits IGroupAddable
    Event Changing As EventHandler(Of CancelableEventArgs(Of String))
    Event Changed As EventHandler(Of EventArgs(Of String))
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
    Inherits ISize
    Inherits IItemDimensions
    Inherits IRowsAndColumns
    Inherits ISelect
    Inherits IItemTemplateable
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IMenuAddable
    Inherits IGroupAddable
    Event SelectionChanging As EventHandler(Of CancelableEventArgs(Of IItem))
    Event SelectionChanged As EventHandler(Of EventArgs(Of IItem))
End Interface

Public Interface IGroupAddable
    Inherits IRibbonElement
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
    Inherits IBoxAddable
    Inherits IGroupAddable
End Interface

Public Interface IMenuAddable
    Inherits IRibbonElement
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
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IMenuAddable
    Inherits IGroupAddable
End Interface

Public Interface IMenuSeparator
    Inherits IRibbonElement
    Inherits ITitle
    Inherits IMenuAddable
End Interface

Public Interface IRibbon
    Inherits IRibbonUI
    Inherits IReadOnlyCollection(Of IRibbonElement)
    Sub ReceiveIRibbonUI(ribbonUi As IRibbonUI)
    Function GetElement(id As String) As IRibbonElement
    Function GetElementProperty(Of TProperty As IRibbonElementProperty)(id As String) As TProperty
    Function GetContainer(Of TRibbonElement As IRibbonElement)(id As String) As IReadOnlyCollection(Of TRibbonElement)
    Function GetContainerItem(parentId As String, index As Integer) As IItem
    ReadOnly Property RibbonX As String
    Function GetErrors() As RibbonErrorLog
End Interface

Public Interface ISeparator
    Inherits IRibbonElement
    Inherits IVisible
    Inherits IButtonGroupAddable
    Inherits IGroupAddable
End Interface

Public Interface ISplitButton
    Inherits IRibbonElement
    Inherits IReadOnlyCollection(Of IRibbonElement)
    Inherits IEnabled
    Inherits IVisible
    Inherits IKeyTip
    Inherits ISize
    Inherits IShowLabel
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IMenuAddable
    Inherits IGroupAddable

    ReadOnly Property Button As IRibbonElement
    ReadOnly Property Menu As IMenu
End Interface

Public Interface ITab
    Inherits IRibbonElement
    Inherits IReadOnlyCollection(Of IGroup)
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
    Inherits IBoxAddable
    Inherits IButtonGroupAddable
    Inherits IMenuAddable
    Inherits IGroupAddable

    Event Changing As EventHandler(Of CancelableEventArgs(Of Boolean))
    Event Changed As EventHandler(Of EventArgs(Of Boolean))
End Interface
Imports RibbonFactory.Controls.Events

Namespace BuilderInterfaces.API

    Public Interface IDropdownBuilder
        Inherits IID(Of IDropdownBuilder)
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
        Inherits IGetSelectedItemId(Of IDropdownBuilder)
        Inherits IGetSelectedItemIndex(Of IDropdownBuilder)

        Function RouteSelectionChangeTo(callBack As SelectionChanged) As IDropdownBuilder
        Function BeforeSelectionChange(ParamArray actions() As Action(Of BeforeSelectionChangeEventArgs)) As IDropdownBuilder
        Function AfterSelectionChanged(ParamArray actions() As Action(Of SelectionChangeEventArgs)) As IDropdownBuilder

    End Interface

End Namespace

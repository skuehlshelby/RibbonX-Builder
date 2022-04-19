Imports RibbonFactory.Controls.Events

Namespace BuilderInterfaces.API

    Public Interface IGalleryBuilder
        Inherits IID(Of IGalleryBuilder)
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
        Inherits IGetSelectedItemId(Of IGalleryBuilder)
        Inherits IGetSelectedItemIndex(Of IGalleryBuilder)

        Function RouteSelectionChangeTo(callback As SelectionChanged) As IGalleryBuilder
        Function BeforeSelectionChange(ParamArray actions() As Action(Of BeforeSelectionChangeEventArgs)) As IGalleryBuilder
        Function AfterSelectionChanged(ParamArray actions() As Action(Of SelectionChangeEventArgs)) As IGalleryBuilder
    End Interface

End Namespace
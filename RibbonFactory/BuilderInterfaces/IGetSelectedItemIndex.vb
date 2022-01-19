Namespace BuilderInterfaces

    Public Interface IGetSelectedItemIndex(Of TBuilder, TRibbonElement As RibbonElement)

        Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As TBuilder

        Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged, onSelectionChange As Action(Of TRibbonElement)) As TBuilder

    End Interface

End Namespace
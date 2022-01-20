Namespace BuilderInterfaces

    Public Interface IGetSelectedItemIndex(Of TBuilder)

        Function GetSelectedItemIndexFrom(getSelected As FromControl(Of Integer), setSelected As SelectionChanged) As TBuilder

    End Interface

End Namespace
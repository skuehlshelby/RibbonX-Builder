Namespace BuilderInterfaces

    Public Interface IGetSelectedItemId(Of T)

        Function GetSelectedItemIdFrom(callback As FromControl(Of String), setSelected As SelectionChanged) As T

    End Interface

End Namespace
Namespace BuilderInterfaces

    Public Interface IGetItemId(Of T)

        Function GetItemIdFrom(callback As FromControlAndIndex(Of String)) As T

    End Interface

End Namespace
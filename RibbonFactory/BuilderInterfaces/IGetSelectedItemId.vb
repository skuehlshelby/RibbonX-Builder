Namespace BuilderInterfaces

    Public Interface IGetSelectedItemId(Of T)

        Function GetSelectedItemIdFrom(callback As FromControl(Of String)) As T

    End Interface

End Namespace
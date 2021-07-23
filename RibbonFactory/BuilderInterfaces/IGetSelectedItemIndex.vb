Namespace BuilderInterfaces

    Public Interface IGetSelectedItemIndex(Of T)

        Function GetSelectedItemIndexFrom(callback As FromControl(Of Integer)) As T

    End Interface

End Namespace
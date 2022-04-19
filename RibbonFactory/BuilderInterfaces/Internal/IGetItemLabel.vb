Namespace BuilderInterfaces

    Public Interface IGetItemLabel(Of T)

        Function GetItemLabelFrom(callback As FromControlAndIndex(Of String)) As T

    End Interface

End Namespace
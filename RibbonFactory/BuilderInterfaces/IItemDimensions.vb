Namespace BuilderInterfaces

    Public Interface IItemDimensions(Of T)
    
        Function WithItemHeight(height As Integer) As T

        Function WithItemHeight(height As Integer, callback As FromControl(Of Integer)) As T

        Function WithItemWidth(width As Integer) As T

        Function WithItemWidth(width As Integer, callback As FromControl(Of Integer)) As T

    End Interface

End NameSpace
Namespace Builders.Internal.Standardization

    Public Interface IRowsAndColumns(Of T)
    
        Function WithRowCount(rowCount As Integer) As T

        Function WithColumnCount(columnCount As Integer) As T

    End Interface

End NameSpace
Friend Module Contracts

    Public Sub Require(Of T As Exception)(condition As Boolean)
        If Not condition Then
            Dim ex As T = Activator.CreateInstance(Of T)
            Throw ex
        End If
    End Sub

    Public Sub Require(Of T As Exception)(condition As Boolean, message As String)
        If Not condition Then
            Dim ex As T = CType(Activator.CreateInstance(GetType(T), message), T)
            Throw ex
        End If
    End Sub

End Module
Namespace Utilities
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

        Public Function NotNull(Of T)(value As T) As T
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value), $"Parameter of type '{GetType(T).Name}' is not allowed to be null here.")
            Else
                Return value
            End If
        End Function

        Public Function NotNullOrWhiteSpace(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then
                Throw New ArgumentNullException(NameOf(value), "String parameter is not allowed to be null or whitespace here.")
            Else
                Return value
            End If
        End Function

    End Module
End NameSpace
Friend Class NullGuard
    'TODO Replace with Contract.Requires
    Public Shared Function NotNull(Of T As Class)(arg As T, argName As String) As T
        If IsNothing(arg) Then
            Throw New ArgumentNullException($"{argName} cannot be null!")
        Else
            Return arg
        End If
    End Function
End Class

Public NotInheritable Class RefreshNeededEventArgs
    Inherits EventArgs

    Public Sub New(id As String)
        Me.ID = id
    End Sub

    Public ReadOnly Property ID As String

End Class

Public Class CancelableEventArgs
    Inherits EventArgs

    Private cancelled As Boolean = False

    Public Sub Cancel()
        If Not cancelled Then
            cancelled = True
        End If
    End Sub

    Public ReadOnly Property IsCancelled As Boolean
        Get
            Return cancelled
        End Get
    End Property

End Class

Public Class CancelableEventArgs(Of T)
    Inherits CancelableEventArgs

    Public Sub New(current As T, future As T)
        Me.Current = current
        Me.Future = future
    End Sub

    Public ReadOnly Property Current As T
    Public ReadOnly Property Future As T
End Class

Public Class EventArgs(Of T)
    Inherits EventArgs

    Public Sub New(current As T)
        Me.Current = current
    End Sub

    Public ReadOnly Property Current As T

End Class

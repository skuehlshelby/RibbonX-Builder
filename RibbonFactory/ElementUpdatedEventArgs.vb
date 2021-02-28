Public Class ElementUpdatedEventArgs
    Inherits EventArgs

    Public Sub New(id As String)
        Me.ID = id
    End Sub

    Public ReadOnly Property ID As String
End Class
Namespace Controls.Events

    Public Class ToggleChangeEventArgs
        Inherits EventArgs

        Public Sub New(currentState As Boolean)
            Me.CurrentState = currentState
        End Sub

        Public ReadOnly Property CurrentState As Boolean

    End Class

End Namespace
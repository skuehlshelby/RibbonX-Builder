Namespace Controls.EventArgs

    Public Class BeforeToggleChangeEventArgs
        Inherits CancelableEventArgs

        Public Sub New(currentState As Boolean)
            Me.CurrentState = currentState
            FutureState = Not currentState
        End Sub

        Public ReadOnly Property CurrentState As Boolean

        Public ReadOnly Property FutureState As Boolean

    End Class

End Namespace

Namespace Controls.Events

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

End Namespace

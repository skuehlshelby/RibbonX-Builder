Imports RibbonX.Utilities

Namespace Controls.EventArgs

    Public Class BeforeTextChangeEventArgs
        Inherits CancelableEventArgs

        Public Sub New(originalText As String, newText As String)
            Me.OriginalText = NonNull(originalText)
            Me.NewText = NonNull(newText)
        End Sub

        Public ReadOnly Property OriginalText As String

        Public Property NewText As String

    End Class

End Namespace
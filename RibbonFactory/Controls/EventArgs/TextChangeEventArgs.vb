Imports RibbonX.Utilities

Namespace Controls.EventArgs

    Public Class TextChangeEventArgs
        Inherits System.EventArgs

        Public Sub New(newText As String)
            Me.NewText = NonNull(newText)
        End Sub

        Public ReadOnly Property NewText As String

    End Class

End Namespace
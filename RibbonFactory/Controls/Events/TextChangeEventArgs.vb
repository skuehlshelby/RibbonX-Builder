Imports RibbonFactory.Utilities

Namespace Controls.Events

    Public Class TextChangeEventArgs
        Inherits EventArgs

        Public Sub New(newText As String)
            Me.NewText = NonNull(newText)
        End Sub

        Public ReadOnly Property NewText As String

    End Class

End Namespace
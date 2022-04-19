Imports RibbonFactory.Utilities

Namespace Controls.Events

    Public Class SelectionChangeEventArgs
        Inherits EventArgs

        Public Sub New(newSelection As Item)
            Me.NewSelection = NonNull(newSelection)
        End Sub

        Public ReadOnly Property NewSelection As Item

    End Class

End Namespace

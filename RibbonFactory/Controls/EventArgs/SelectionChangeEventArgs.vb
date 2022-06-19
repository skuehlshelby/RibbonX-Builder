Imports RibbonX.Utilities

Namespace Controls.EventArgs

    Public Class SelectionChangeEventArgs
        Inherits System.EventArgs

        Public Sub New(newSelection As Item)
            Me.NewSelection = NonNull(newSelection)
        End Sub

        Public ReadOnly Property NewSelection As Item

    End Class

End Namespace

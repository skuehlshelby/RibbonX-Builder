Imports RibbonX.Utilities

Namespace Controls.Events

    Public Class BeforeSelectionChangeEventArgs
        Inherits CancelableEventArgs

        Public Sub New(availableSelections As ICollection(Of Item), originalSelection As Item, newSelection As Item)
            Me.AvailableSelections = NonNull(availableSelections)
            Me.OriginalSelection = originalSelection
            Me.NewSelection = newSelection
        End Sub

        Public ReadOnly Property AvailableSelections As ICollection(Of Item)

        Public ReadOnly Property OriginalSelection As Item

        Public Property NewSelection As Item

    End Class

End Namespace

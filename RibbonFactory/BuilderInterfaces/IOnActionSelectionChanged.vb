Namespace BuilderInterfaces

    Public Interface IOnActionSelectionChanged(Of T)

        Function ThatDoes(action As Action, callback As SelectionChanged) As T

    End Interface
End Namespace
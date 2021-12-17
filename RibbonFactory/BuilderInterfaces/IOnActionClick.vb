Namespace BuilderInterfaces

    Public Interface IOnActionClick(Of T)

        Function ThatDoes(action As Action, callback As OnAction) As T

    End Interface

End NameSpace
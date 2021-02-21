Namespace Builder_Interfaces
    Public Interface IActionable(Of Out T)
        Function ThatDoes(Action As System.Action) As T
    End Interface
End NameSpace
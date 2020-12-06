Public Interface IActionable(Of Out T)
    Function ThatDoes(ByVal Action As System.Action) As T
End Interface

Public Interface IVeryDescribable(Of Out T)
    Overloads Function WithDescription(ByVal Description As String) As T
    Overloads Function WithDescription(ByVal Description As String, ByVal Callback As CallbackSignatures.FromControl(Of String)) As T
End Interface

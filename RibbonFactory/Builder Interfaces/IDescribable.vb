Public Interface IDescribable(Of Out T)
    Overloads Function WithLabel(ByVal Label As String, Optional ByVal CopyToScreentip As Boolean = True) As T
    Overloads Function WithLabel(ByVal Label As String, ByVal Callback As CallbackSignatures.FromControl(Of String), Optional ByVal CopyToScreentip As Boolean = True) As T
    Overloads Function WithScreenTip(ByVal ScreenTip As String) As T
    Overloads Function WithScreenTip(ByVal ScreenTip As String, ByVal Callback As CallbackSignatures.FromControl(Of String)) As T
    Overloads Function WithSupertip(ByVal Supertip As String) As T
    Overloads Function WithSupertip(ByVal Supertip As String, ByVal Callback As CallbackSignatures.FromControl(Of String)) As T
End Interface